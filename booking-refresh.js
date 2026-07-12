(function () {
  "use strict";

  const BOOKING_URL =
    "https://raw.githubusercontent.com/BOBWZW2/data-base/main/booking_latest.xlsx";
  const BOOKING_COMMITS_API =
    "https://api.github.com/repos/BOBWZW2/data-base/commits?path=booking_latest.xlsx&per_page=5";
  const KNOWN_GOOD_FALLBACKS = [
    {
      sha: "eb8542c84d07db8664c27c3a6c37ebc7f55e03b8",
      committedAt: "2026-07-09T01:19:41Z",
    },
  ];
  const DATABASE_NAME = "ts-booking-control";
  const DATABASE_VERSION = 1;
  const STORE_NAME = "booking-snapshots";
  const CACHE_KEY = "latest";
  const CACHE_SCHEMA_VERSION = 1;

  let xlsxModulePromise = null;

  function singaporeDateKey(date = new Date()) {
    const parts = new Intl.DateTimeFormat("en-CA", {
      timeZone: "Asia/Singapore",
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
    }).formatToParts(date);
    const values = Object.fromEntries(parts.map((part) => [part.type, part.value]));
    return `${values.year}-${values.month}-${values.day}`;
  }

  function openDatabase() {
    if (!("indexedDB" in window)) return Promise.resolve(null);
    return new Promise((resolve, reject) => {
      const request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION);
      request.onupgradeneeded = () => {
        const database = request.result;
        if (!database.objectStoreNames.contains(STORE_NAME)) {
          database.createObjectStore(STORE_NAME);
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("无法打开浏览器数据缓存"));
    });
  }

  async function readCachedSnapshot() {
    let database;
    try {
      database = await openDatabase();
      if (!database) return null;
      return await new Promise((resolve, reject) => {
        const transaction = database.transaction(STORE_NAME, "readonly");
        const request = transaction.objectStore(STORE_NAME).get(CACHE_KEY);
        request.onsuccess = () => resolve(request.result || null);
        request.onerror = () => reject(request.error || new Error("无法读取浏览器数据缓存"));
      });
    } catch (error) {
      console.warn("Booking cache read failed", error);
      return null;
    } finally {
      database?.close();
    }
  }

  async function writeCachedSnapshot(snapshot) {
    let database;
    try {
      database = await openDatabase();
      if (!database) return false;
      await new Promise((resolve, reject) => {
        const transaction = database.transaction(STORE_NAME, "readwrite");
        transaction.objectStore(STORE_NAME).put(snapshot, CACHE_KEY);
        transaction.oncomplete = () => resolve();
        transaction.onerror = () =>
          reject(transaction.error || new Error("无法保存浏览器数据缓存"));
        transaction.onabort = () =>
          reject(transaction.error || new Error("浏览器数据缓存写入已取消"));
      });
      return true;
    } catch (error) {
      console.warn("Booking cache write failed", error);
      return false;
    } finally {
      database?.close();
    }
  }

  async function loadXlsxLibrary() {
    if (window.XLSX?.utils && typeof window.XLSX.read === "function") return window.XLSX;
    if (!xlsxModulePromise) {
      xlsxModulePromise = import("./vendor/xlsx.mjs?v=0.20.3").then((module) => {
        if (!module?.utils || typeof module.read !== "function") {
          throw new Error("Excel 解析组件载入失败");
        }
        return module;
      });
    }
    return xlsxModulePromise;
  }

  function cleanText(value) {
    if (value === null || value === undefined) return "";
    const result = String(value).trim();
    if (["nan", "none"].includes(result.toLowerCase())) return "";
    return result.replace(/\s+/g, " ");
  }

  function normalizedKey(value) {
    return cleanText(value).toUpperCase();
  }

  function numericValue(value) {
    if (typeof value === "number") return Number.isFinite(value) ? value : null;
    const cleaned = cleanText(value).replaceAll(",", "");
    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : null;
  }

  function valueOrZero(value) {
    return numericValue(value) ?? 0;
  }

  function roundedQuantity(value) {
    const rounded = Math.round(Number(value || 0) * 1000) / 1000;
    if (Math.abs(rounded) < 0.0005) return 0;
    return rounded;
  }

  function adjustedForty(count20, count40, teu) {
    if (teu === null) return count40;
    const adjusted = (teu - count20) / 2;
    return adjusted < -0.001 ? count40 : adjusted;
  }

  function stage(rowValue, prefix, fallback = {}) {
    let count20 = numericValue(rowValue(`${prefix} Unit 20ft`));
    let count40 = numericValue(rowValue(`${prefix} Unit 40ft`));
    let teu = numericValue(rowValue(`${prefix} TTL Teu`));
    if (count20 === null) count20 = fallback.t20 ? valueOrZero(rowValue(fallback.t20)) : 0;
    if (count40 === null) count40 = fallback.t40 ? valueOrZero(rowValue(fallback.t40)) : 0;
    if (teu === null && fallback.teu) teu = numericValue(rowValue(fallback.teu));
    if (teu === null) teu = count20 + 2 * count40;
    count40 = adjustedForty(count20, count40, teu);
    return {
      t20: roundedQuantity(count20),
      t40: roundedQuantity(count40),
      teu: roundedQuantity(teu),
    };
  }

  function maxStage(left, right) {
    return {
      t20: roundedQuantity(Math.max(left.t20, right.t20)),
      t40: roundedQuantity(Math.max(left.t40, right.t40)),
      teu: roundedQuantity(Math.max(left.teu, right.teu)),
    };
  }

  function increment(counter, key) {
    if (!key) return;
    counter.set(key, (counter.get(key) || 0) + 1);
  }

  function incrementNested(counter, firstKey, secondKey) {
    if (!firstKey || !secondKey) return;
    if (!counter.has(firstKey)) counter.set(firstKey, new Map());
    increment(counter.get(firstKey), secondKey);
  }

  function sortedKeys(counter) {
    return [...counter.entries()]
      .sort((left, right) => right[1] - left[1] || left[0].localeCompare(right[0], "en"))
      .map(([key]) => key);
  }

  function legColumn(index, name) {
    const suffix = index === 1 ? "st" : index === 2 ? "nd" : index === 3 ? "rd" : "th";
    return `${index}${suffix} ${name}`;
  }

  function buildBookingSnapshot(arrayBuffer) {
    return loadXlsxLibrary().then((XLSX) => {
      const workbook = XLSX.read(arrayBuffer, { type: "array", cellDates: false });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error("Excel 中没有可读取的工作表");
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
        header: 1,
        defval: "",
        raw: true,
      });
      if (rows.length < 2) throw new Error("Excel 中没有 Booking 数据");

      const headerIndexes = new Map(
        rows[0].map((header, index) => [cleanText(header), index]),
      );
      for (const required of ["1st VVD", "Booking Unit 20ft", "Booking Unit 40ft"]) {
        if (!headerIndexes.has(required)) {
          throw new Error(`Excel 缺少必要字段：${required}`);
        }
      }

      const records = [];
      const allLanes = new Map();
      const allVvds = new Map();
      const allPorts = new Map();
      const laneVvdCounts = new Map();
      const vvdLaneCounts = new Map();

      for (let rowIndex = 1; rowIndex < rows.length; rowIndex += 1) {
        const row = rows[rowIndex];
        if (!row?.length) continue;
        const rowValue = (name) => {
          const columnIndex = headerIndexes.get(name);
          return columnIndex === undefined ? "" : row[columnIndex];
        };

        const booking = stage(rowValue, "Booking", {
          t20: "20ft",
          t40: "40ft",
          teu: "Booking",
        });
        const vl = stage(rowValue, "VL");
        const op = maxStage(stage(rowValue, "OP"), vl);
        const weightTon = valueOrZero(rowValue("Container Weight")) / 1000 + booking.teu * 2;
        const legs = [];

        for (let legIndex = 1; legIndex <= 5; legIndex += 1) {
          const pol = cleanText(rowValue(legColumn(legIndex, "POL")));
          const pod = cleanText(rowValue(legColumn(legIndex, "POD")));
          const lane = cleanText(rowValue(legColumn(legIndex, "Lane")));
          const vvd = cleanText(rowValue(legColumn(legIndex, "VVD")));
          if (!(pol || pod || lane || vvd)) continue;
          legs.push([pol, pod, lane, vvd]);
          increment(allLanes, lane);
          increment(allVvds, vvd);
          increment(allPorts, pol);
          increment(allPorts, pod);
          incrementNested(laneVvdCounts, lane, vvd);
          incrementNested(vvdLaneCounts, vvd, lane);
        }
        if (!legs.length) continue;

        records.push([
          normalizedKey(rowValue("COC/SOC")),
          normalizedKey(rowValue("SUL YN")) === "Y" ? "SUL" : "CUL",
          booking.t20,
          booking.t40,
          booking.teu,
          op.t20,
          op.t40,
          op.teu,
          vl.t20,
          vl.t40,
          vl.teu,
          roundedQuantity(weightTon),
          legs,
          cleanText(rowValue("BL No.")),
          cleanText(rowValue("CUL CODE")),
          cleanText(rowValue("POR")),
          cleanText(rowValue("POL")),
          cleanText(rowValue("POD")),
          cleanText(rowValue("DEL")),
        ]);
      }

      const laneVvds = Object.fromEntries(
        [...laneVvdCounts.entries()]
          .sort(([left], [right]) => left.localeCompare(right, "en"))
          .map(([lane, counter]) => [lane, sortedKeys(counter)]),
      );
      const vvdLanes = Object.fromEntries(
        [...vvdLaneCounts.entries()]
          .sort(([left], [right]) => left.localeCompare(right, "en"))
          .map(([vvd, counter]) => [
            vvd,
            [...counter.entries()]
              .sort(
                (left, right) =>
                  right[1] - left[1] || left[0].localeCompare(right[0], "en"),
              )
              .map(([lane, count]) => ({ lane, count })),
          ]),
      );

      const refreshedAt = new Date().toISOString();
      return {
        cacheSchemaVersion: CACHE_SCHEMA_VERSION,
        dateKey: singaporeDateKey(),
        refreshedAt,
        sourceUrl: BOOKING_URL,
        booking: {
          meta: {
            source: "BOBWZW2/data-base booking_latest.xlsx",
            generatedRows: records.length,
            xlsxSize: arrayBuffer.byteLength,
            refreshedAt,
            refreshDate: singaporeDateKey(),
          },
          lanes: sortedKeys(allLanes),
          vvds: sortedKeys(allVvds),
          ports: sortedKeys(allPorts),
          laneVvds,
          vvdLanes,
          records,
        },
      };
    });
  }

  function mergeSnapshot(basePayload, snapshot) {
    if (!snapshot?.booking || snapshot.cacheSchemaVersion !== CACHE_SCHEMA_VERSION) {
      return basePayload;
    }
    const booking = snapshot.booking;
    return {
      ...basePayload,
      meta: { ...basePayload.meta, ...booking.meta },
      lanes: booking.lanes,
      vvds: booking.vvds,
      ports: booking.ports,
      laneVvds: booking.laneVvds,
      vvdLanes: booking.vvdLanes,
      records: booking.records,
    };
  }

  async function downloadLatestSnapshot(onProgress) {
    onProgress?.("正在从 GitHub 下载 Daily Booking Excel…");
    const separator = BOOKING_URL.includes("?") ? "&" : "?";
    const latestUrl = `${BOOKING_URL}${separator}refresh=${Date.now()}`;
    let snapshot;
    let latestError;

    try {
      snapshot = await downloadAndBuild(latestUrl, onProgress);
    } catch (error) {
      latestError = error;
      console.warn("Latest Booking workbook is unavailable", error);
    }

    if (!snapshot) {
      onProgress?.("最新版 Excel 不可用，正在查找最近可用版本…");
      let commits = [];
      try {
        const response = await fetch(`${BOOKING_COMMITS_API}&refresh=${Date.now()}`, {
          cache: "no-store",
          headers: { Accept: "application/vnd.github+json" },
        });
        if (response.ok) commits = await response.json();
      } catch (error) {
        console.warn("Booking commit history lookup failed", error);
      }

      const historicalCandidates = [
        ...commits.slice(1, 5).map((commit) => ({
          sha: commit?.sha,
          committedAt: commit?.commit?.committer?.date || commit?.commit?.author?.date,
        })),
        ...KNOWN_GOOD_FALLBACKS,
      ].filter(
        (candidate, index, candidates) =>
          candidate?.sha &&
          candidates.findIndex((item) => item?.sha === candidate.sha) === index,
      );

      for (const candidate of historicalCandidates) {
        const commitSha = cleanText(candidate?.sha);
        if (!commitSha) continue;
        const commitAt = cleanText(candidate?.committedAt);
        onProgress?.(
          `正在尝试 GitHub 最近可用版本${commitAt ? `（${commitAt.slice(0, 10)}）` : ""}…`,
        );
        try {
          const historicalUrl = `https://raw.githubusercontent.com/BOBWZW2/data-base/${commitSha}/booking_latest.xlsx`;
          snapshot = await downloadAndBuild(historicalUrl, onProgress);
          snapshot.fallbackSource = true;
          snapshot.sourceCommit = commitSha;
          snapshot.sourceCommitAt = commitAt;
          snapshot.booking.meta.fallbackSource = true;
          snapshot.booking.meta.sourceCommit = commitSha;
          snapshot.booking.meta.sourceCommitAt = commitAt;
          break;
        } catch (error) {
          console.warn(`Booking workbook ${commitSha} is unavailable`, error);
        }
      }
    }

    if (!snapshot) throw latestError || new Error("GitHub 上没有可读取的 Booking Excel");
    onProgress?.("正在保存今日数据…");
    await writeCachedSnapshot(snapshot);
    return snapshot;
  }

  async function downloadAndBuild(url, onProgress) {
    const response = await fetch(url, {
      cache: "no-store",
      headers: { Accept: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
    });
    if (!response.ok) throw new Error(`GitHub 下载失败（HTTP ${response.status}）`);
    const arrayBuffer = await response.arrayBuffer();
    if (arrayBuffer.byteLength < 100000) throw new Error("GitHub 返回的 Excel 文件不完整");
    onProgress?.("Excel 已下载，正在整理 Booking 数据…");
    return buildBookingSnapshot(arrayBuffer);
  }

  window.TSBookingRefresh = {
    BOOKING_URL,
    CACHE_SCHEMA_VERSION,
    singaporeDateKey,
    readCachedSnapshot,
    downloadLatestSnapshot,
    mergeSnapshot,
    buildBookingSnapshot,
  };
})();
