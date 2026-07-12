const state = {
  data: null,
  recordsByVvd: new Map(),
  voyageCache: new Map(),
  activeModule: "direct",
  expandedPols: new Set(),
  expandedConnections: new Set(),
  transferDirection: "all",
  transferRisk: "all",
  transferType: "all",
  selectedTransferPort: "",
  selectedTransferSegment: "",
  expandedFeederGroups: new Set(),
  drawerSegmentKey: "",
  xlsxModulePromise: null,
  lastQuerySignature: "",
  basePayload: null,
  refreshingData: false,
  transferEvents: [],
  transferEventsByPort: new Map(),
  transferSearchSubmitted: false,
  transferSearchFilters: Object.fromEntries(
    ["from", "to", "bookingPol", "bookingPod", "laneIn", "laneOut", "vvdIn", "vvdOut"].map((field) => [field, new Set()]),
  ),
  transferCallSelections: new Map(),
  transferSearchGroupEvents: new Map(),
  transferMissingGroupEvents: new Map(),
};

const els = {
  meta: document.getElementById("dataMeta"),
  dataStatus: document.getElementById("dataStatus"),
  refreshBookingButton: document.getElementById("refreshBookingButton"),
  laneInput: document.getElementById("laneInput"),
  vvdInput: document.getElementById("vvdInput"),
  cocFilter: document.getElementById("cocFilter"),
  sulFilter: document.getElementById("sulFilter"),
  laneList: document.getElementById("laneList"),
  vvdList: document.getElementById("vvdList"),
  queryHint: document.getElementById("queryHint"),
  emptyState: document.getElementById("emptyState"),
  emptyTitle: document.getElementById("emptyTitle"),
  emptyText: document.getElementById("emptyText"),
  resultContent: document.getElementById("resultContent"),
  queryLane: document.getElementById("queryLane"),
  overviewTitle: document.getElementById("overviewTitle"),
  querySummary: document.getElementById("querySummary"),
  matchTeu: document.getElementById("matchTeu"),
  matchWeight: document.getElementById("matchWeight"),
  directTeu: document.getElementById("directTeu"),
  directWeight: document.getElementById("directWeight"),
  tsTeu: document.getElementById("tsTeu"),
  tsWeight: document.getElementById("tsWeight"),
  moduleSwitch: document.getElementById("moduleSwitch"),
  voyageSearchWorkbench: document.getElementById("voyageSearchWorkbench"),
  transferSearchWorkbench: document.getElementById("transferSearchWorkbench"),
  transferPortInput: document.getElementById("transferPortInput"),
  transferPortList: document.getElementById("transferPortList"),
  transferStartDate: document.getElementById("transferStartDate"),
  transferEndDate: document.getElementById("transferEndDate"),
  transferSearchButton: document.getElementById("transferSearchButton"),
  transferSearchHint: document.getElementById("transferSearchHint"),
  clearTransferSearchButton: document.getElementById("clearTransferSearchButton"),
  clearTransferOptionalButton: document.getElementById("clearTransferOptionalButton"),
  transferMultiFilters: document.getElementById("transferMultiFilters"),
  transferCocFilter: document.getElementById("transferCocFilter"),
  transferSulFilter: document.getElementById("transferSulFilter"),
  transferSearchContent: document.getElementById("transferSearchContent"),
  transferSearchSwitchTeu: document.getElementById("transferSearchSwitchTeu"),
  transferSearchSwitchWeight: document.getElementById("transferSearchSwitchWeight"),
  transferSearchSummaryTitle: document.getElementById("transferSearchSummaryTitle"),
  transferSearchSummaryHint: document.getElementById("transferSearchSummaryHint"),
  transferSearchMatchedCount: document.getElementById("transferSearchMatchedCount"),
  transferSearchPendingCount: document.getElementById("transferSearchPendingCount"),
  transferSearchMissingCount: document.getElementById("transferSearchMissingCount"),
  transferTotalB20: document.getElementById("transferTotalB20"),
  transferTotalB40: document.getElementById("transferTotalB40"),
  transferTotalBTeu: document.getElementById("transferTotalBTeu"),
  transferTotalOp20: document.getElementById("transferTotalOp20"),
  transferTotalOp40: document.getElementById("transferTotalOp40"),
  transferTotalOpTeu: document.getElementById("transferTotalOpTeu"),
  transferTotalVl20: document.getElementById("transferTotalVl20"),
  transferTotalVl40: document.getElementById("transferTotalVl40"),
  transferTotalVlTeu: document.getElementById("transferTotalVlTeu"),
  transferTotalWeight: document.getElementById("transferTotalWeight"),
  transferCallChoiceSection: document.getElementById("transferCallChoiceSection"),
  transferCallChoices: document.getElementById("transferCallChoices"),
  transferSearchResults: document.getElementById("transferSearchResults"),
  transferMissingSection: document.getElementById("transferMissingSection"),
  transferMissingResults: document.getElementById("transferMissingResults"),
  directSwitchTeu: document.getElementById("directSwitchTeu"),
  directSwitchWeight: document.getElementById("directSwitchWeight"),
  transferSwitchTeu: document.getElementById("transferSwitchTeu"),
  transferSwitchWeight: document.getElementById("transferSwitchWeight"),
  moduleSummaryKicker: document.getElementById("moduleSummaryKicker"),
  moduleSummaryTitle: document.getElementById("moduleSummaryTitle"),
  moduleSummaryHint: document.getElementById("moduleSummaryHint"),
  totalWeight: document.getElementById("totalWeight"),
  totalB20: document.getElementById("totalB20"),
  totalB40: document.getElementById("totalB40"),
  totalBTeu: document.getElementById("totalBTeu"),
  totalOp20: document.getElementById("totalOp20"),
  totalOp40: document.getElementById("totalOp40"),
  totalOpTeu: document.getElementById("totalOpTeu"),
  totalVl20: document.getElementById("totalVl20"),
  totalVl40: document.getElementById("totalVl40"),
  totalVlTeu: document.getElementById("totalVlTeu"),
  directPanel: document.getElementById("directPanel"),
  transferPanel: document.getElementById("transferPanel"),
  directBadge: document.getElementById("directBadge"),
  transferBadge: document.getElementById("transferBadge"),
  directRows: document.getElementById("directRows"),
  transferRows: document.getElementById("transferRows"),
  transferDirection: document.getElementById("transferDirection"),
  transferDirectionHint: document.getElementById("transferDirectionHint"),
  transferPortTabs: document.getElementById("transferPortTabs"),
  clearButton: document.getElementById("clearButton"),
};

const fmt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });
const fmtInt = new Intl.NumberFormat("en-US", { maximumFractionDigits: 0 });
const QUERY_STORAGE_KEY = "tsBookingChecking.lastQuery";

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function normalizedKey(value) {
  return String(value || "").trim().toUpperCase();
}

function joinKey(...parts) {
  return parts.map(normalizedKey).join("|");
}

function number(value) {
  const parsed = Number(value || 0);
  return Number.isFinite(parsed) ? parsed : 0;
}

function display(value, integer = false) {
  return integer ? fmtInt.format(number(value)) : fmt.format(number(value));
}

function arrayValue(value) {
  if (!value) return [];
  return Array.isArray(value) ? value : [value];
}

function optionList(values, limit = 4000) {
  return (values || [])
    .slice(0, limit)
    .map((value) => `<option value="${escapeHtml(value)}"></option>`)
    .join("");
}

function compareAlpha(left, right) {
  return String(left || "").localeCompare(String(right || ""), "en", {
    numeric: true,
    sensitivity: "base",
  });
}

function dateMs(value) {
  if (!value) return Number.NaN;
  return Date.parse(String(value).replace(" ", "T"));
}

function formatDate(value, compact = false) {
  const match = /^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/.exec(String(value || ""));
  if (!match) return value ? String(value) : "暂无";
  return compact
    ? `${match[2]}-${match[3]} ${match[4]}:${match[5]}`
    : `${match[1]}-${match[2]}-${match[3]} ${match[4]}:${match[5]}`;
}

function formatDateOnly(value, compact = false) {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(value || ""));
  if (!match) return value ? String(value) : "暂无";
  return compact ? `${match[2]}-${match[3]}` : `${match[1]}-${match[2]}-${match[3]}`;
}

function formatRefreshTime(value) {
  if (!value) return "";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return new Intl.DateTimeFormat("zh-CN", {
    timeZone: "Asia/Singapore",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    hour12: false,
  }).format(date);
}

function dataVersionLabel(meta = state.data?.meta) {
  const match = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(meta?.sourceCommitAt || ""));
  return match ? `${match[2]}-${match[3]}` : "";
}

function earliest(values) {
  return [...values].filter(Boolean).sort()[0] || "";
}

function normalizeRecord(record, id) {
  if (!Array.isArray(record)) {
    return {
      ...record,
      _id: id,
      legs: (record.legs || []).map((leg) => ({ ...leg })),
    };
  }

  if (record.length >= 13) {
    const [
      coc,
      sul,
      b20,
      b40,
      bTeu,
      op20,
      op40,
      opTeu,
      vl20,
      vl40,
      vlTeu,
      weight,
      legs,
      blNo = "",
      culCode = "",
      por = "",
      pol = "",
      pod = "",
      del = "",
    ] = record;
    return {
      _id: id,
      coc,
      sul,
      b: { t20: b20, t40: b40, teu: bTeu },
      op: { t20: op20, t40: op40, teu: opTeu },
      vl: { t20: vl20, t40: vl40, teu: vlTeu },
      w: weight,
      legs: (legs || []).map(([pol, pod, lane, vvd]) => ({ pol, pod, lane, vvd })),
      blNo,
      culCode,
      original: { por, pol, pod, del },
    };
  }

  const [coc, sul, b20, b40, bTeu, opTeu, vlTeu, weight, legs] = record;
  return {
    _id: id,
    coc,
    sul,
    b: { t20: b20, t40: b40, teu: bTeu },
    op: { t20: 0, t40: 0, teu: opTeu },
    vl: { t20: 0, t40: 0, teu: vlTeu },
    w: weight,
    legs: (legs || []).map(([pol, pod, lane, vvd]) => ({ pol, pod, lane, vvd })),
  };
}

function newTotals() {
  return {
    b20: 0,
    b40: 0,
    bTeu: 0,
    op20: 0,
    op40: 0,
    opTeu: 0,
    vl20: 0,
    vl40: 0,
    vlTeu: 0,
    weight: 0,
  };
}

function addStage(target, source, prefix) {
  target[`${prefix}20`] += number(source?.t20);
  target[`${prefix}40`] += number(source?.t40);
  target[`${prefix}Teu`] += number(source?.teu);
}

function addRecordTotals(target, record) {
  addStage(target, record.b, "b");
  addStage(target, record.op, "op");
  addStage(target, record.vl, "vl");
  target.weight += number(record.w);
}

function addTotals(target, source) {
  for (const key of Object.keys(target)) target[key] += number(source?.[key]);
}

function addRecordOnce(bucket, record) {
  if (bucket.recordIds.has(record._id)) return false;
  bucket.recordIds.add(record._id);
  addRecordTotals(bucket.totals, record);
  return true;
}

function totalsForMatches(matches) {
  const totals = newTotals();
  const seen = new Set();
  for (const { record } of matches) {
    if (seen.has(record._id)) continue;
    seen.add(record._id);
    addRecordTotals(totals, record);
  }
  return { totals, count: seen.size };
}

function filterValue(group) {
  return group.querySelector(".filter-option.active")?.dataset.value || "ALL";
}

function setFilterValue(group, value) {
  const requested = value || "ALL";
  let matched = false;
  group.querySelectorAll(".filter-option").forEach((button) => {
    const active = button.dataset.value === requested;
    matched ||= active;
    button.classList.toggle("active", active);
    button.setAttribute("aria-pressed", active ? "true" : "false");
  });
  if (!matched && requested !== "ALL") setFilterValue(group, "ALL");
}

function recordPassesFilters(record) {
  const coc = filterValue(els.cocFilter);
  const sul = filterValue(els.sulFilter);
  if (coc !== "ALL" && normalizedKey(record.coc) !== coc) return false;
  if (sul !== "ALL" && normalizedKey(record.sul) !== sul) return false;
  return true;
}

function indexRecords() {
  state.recordsByVvd.clear();
  for (const record of state.data.records) {
    record.legs.forEach((leg, legIndex) => {
      const vvd = normalizedKey(leg.vvd);
      if (!vvd) return;
      if (!state.recordsByVvd.has(vvd)) state.recordsByVvd.set(vvd, []);
      state.recordsByVvd.get(vvd).push({ record, legIndex });
    });
  }
}

function collectMatches() {
  const vvd = normalizedKey(els.vvdInput.value);
  const lane = normalizedKey(els.laneInput.value);
  if (!vvd) return [];
  const candidates = state.recordsByVvd.get(vvd) || [];
  const seen = new Set();
  const matches = [];
  for (const match of candidates) {
    const leg = match.record.legs[match.legIndex];
    if (lane && normalizedKey(leg.lane) !== lane) continue;
    if (!recordPassesFilters(match.record)) continue;
    const key = `${match.record._id}|${match.legIndex}`;
    if (seen.has(key)) continue;
    seen.add(key);
    matches.push(match);
  }
  return matches;
}

function parseVoyageCall(raw, lane, vvd) {
  if (!Array.isArray(raw)) return { ...raw, lane, vvd };
  const [port, eta, etb, etd, callSeq, sourceRow] = raw;
  return {
    port: normalizedKey(port),
    eta: eta || "",
    etb: etb || "",
    etd: etd || "",
    callSeq: number(callSeq),
    sourceRow: number(sourceRow),
    lane: normalizedKey(lane),
    vvd: normalizedKey(vvd),
  };
}

function voyageFor(lane, vvd) {
  const key = joinKey(lane, vvd);
  if (state.voyageCache.has(key)) return state.voyageCache.get(key);
  const voyage = (state.data?.schedule?.voyages?.[key] || []).map((raw) =>
    parseVoyageCall(raw, lane, vvd),
  );
  state.voyageCache.set(key, voyage);
  return voyage;
}

function parseScheduleCall(raw, source, leg, port) {
  if (!Array.isArray(raw)) {
    return {
      ...raw,
      lane: normalizedKey(raw.lane || leg.lane),
      vvd: normalizedKey(raw.vvd || leg.vvd),
      port: normalizedKey(raw.port || port),
      source,
    };
  }

  if (source === "fallback" && raw.length >= 6) {
    const [lane, eta, etb, etd, callSeq, sourceRow] = raw;
    return {
      lane: normalizedKey(lane),
      vvd: normalizedKey(leg.vvd),
      port: normalizedKey(port),
      eta: eta || "",
      etb: etb || "",
      etd: etd || "",
      callSeq: number(callSeq),
      sourceRow: number(sourceRow),
      source,
    };
  }

  const [eta, etb, etd, callSeqOrRow, sourceRow] = raw;
  return {
    lane: normalizedKey(leg.lane),
    vvd: normalizedKey(leg.vvd),
    port: normalizedKey(port),
    eta: eta || "",
    etb: etb || "",
    etd: etd || "",
    callSeq: raw.length >= 5 ? number(callSeqOrRow) : 0,
    sourceRow: raw.length >= 5 ? number(sourceRow) : number(callSeqOrRow),
    source,
  };
}

function legacyScheduleCalls(leg, port) {
  const schedule = state.data?.schedule;
  const exact = schedule?.byLeg?.[joinKey(leg.lane, leg.vvd, port)];
  const fallback = schedule?.byVvdPort?.[joinKey(leg.vvd, port)];
  return arrayValue(exact || fallback).map((etd) => ({
    lane: normalizedKey(leg.lane),
    vvd: normalizedKey(leg.vvd),
    port: normalizedKey(port),
    eta: "",
    etb: "",
    etd,
    callSeq: 0,
    sourceRow: 0,
    source: exact ? "legacy-exact" : "legacy-fallback",
  }));
}

function routeOverrideCalls(leg, port) {
  const schedule = state.data?.schedule;
  const routeKey = joinKey(leg.lane, leg.vvd, leg.pol, leg.pod);
  const override = schedule?.routeOverrides?.[routeKey];
  if (!override) return [];

  const normalizedPort = normalizedKey(port);
  let rawCalls = [];
  if (normalizedPort === normalizedKey(leg.pol) && override[0]?.length) {
    rawCalls = override[0];
  } else if (normalizedPort === normalizedKey(leg.pod) && override[1]?.length) {
    rawCalls = override[1];
  }

  return rawCalls.map(([actualVvd, eta, etb, etd, callSeq, sourceRow, actualLane]) => ({
    lane: normalizedKey(actualLane || leg.lane),
    vvd: normalizedKey(actualVvd),
    port: normalizedPort,
    eta: eta || "",
    etb: etb || "",
    etd: etd || "",
    callSeq: number(callSeq),
    sourceRow: number(sourceRow),
    source: "route-override",
    resolvedFromLane: normalizedKey(leg.lane),
    resolvedFromVvd: normalizedKey(leg.vvd),
    preferred: true,
  }));
}

function scheduleCallsFor(leg, port) {
  const schedule = state.data?.schedule;
  if (!schedule || !leg || !normalizedKey(leg.vvd) || !normalizedKey(port)) return [];

  const overrideCalls = routeOverrideCalls(leg, port);
  if (overrideCalls.length) return overrideCalls;

  const exactRaw = schedule.callsByLeg?.[joinKey(leg.lane, leg.vvd, port)] || [];
  if (exactRaw.length) {
    return exactRaw.map((raw) => parseScheduleCall(raw, "exact", leg, port));
  }

  const fallbackRaw = schedule.callsByVvdPort?.[joinKey(leg.vvd, port)] || [];
  if (fallbackRaw.length) {
    const calls = fallbackRaw.map((raw) => parseScheduleCall(raw, "fallback", leg, port));
    const lane = normalizedKey(leg.lane);
    const sameLane = lane ? calls.filter((call) => call.lane === lane) : [];
    return lane ? sameLane : calls;
  }

  return legacyScheduleCalls(leg, port);
}

function voyageIndexForCall(call) {
  const voyage = voyageFor(call.lane, call.vvd);
  if (!voyage.length) return { voyage, index: -1 };
  let index = call.callSeq
    ? voyage.findIndex((item) => item.callSeq === call.callSeq)
    : -1;
  if (index < 0 && call.sourceRow) {
    index = voyage.findIndex((item) => item.sourceRow === call.sourceRow);
  }
  if (index < 0) {
    index = voyage.findIndex(
      (item) =>
        item.port === call.port &&
        item.eta === call.eta &&
        item.etb === call.etb &&
        item.etd === call.etd,
    );
  }
  return { voyage, index };
}

function relativePortDistance(call, direction, targetPort) {
  const target = normalizedKey(targetPort);
  if (!target) return null;
  const { voyage, index } = voyageIndexForCall(call);
  if (index < 0) return null;
  const step = direction === "before" ? -1 : 1;
  for (let cursor = index + step; cursor >= 0 && cursor < voyage.length; cursor += step) {
    if (voyage[cursor].port === target) return Math.abs(cursor - index);
  }
  return null;
}

function selectScheduleCalls(leg, port) {
  const rawCalls = scheduleCallsFor(leg, port);
  if (!rawCalls.length) {
    return { calls: [], rawCalls: [], rawCount: 0, narrowed: false };
  }

  const normalizedPort = normalizedKey(port);
  const requirements = [];
  if (normalizedPort === normalizedKey(leg.pol) && normalizedKey(leg.pod)) {
    requirements.push({ direction: "after", target: leg.pod });
  }
  if (normalizedPort === normalizedKey(leg.pod) && normalizedKey(leg.pol)) {
    requirements.push({ direction: "before", target: leg.pol });
  }

  let selected = rawCalls;
  if (requirements.length) {
    const scored = rawCalls
      .map((call) => {
        const distances = requirements.map((requirement) =>
          relativePortDistance(call, requirement.direction, requirement.target),
        );
        return {
          call,
          valid: distances.every((distance) => distance !== null),
          score: distances.reduce((sum, distance) => sum + number(distance), 0),
        };
      })
      .filter((item) => item.valid);
    if (scored.length) {
      const minimum = Math.min(...scored.map((item) => item.score));
      selected = scored.filter((item) => item.score === minimum).map((item) => item.call);
    }
  }

  const preferred = selected.filter((call) => call.preferred);
  if (preferred.length) selected = preferred;

  const sortValue = (call) => call.etb || call.etd || call.eta || "9999";
  selected = [...selected].sort((left, right) => {
    const byTime = sortValue(left).localeCompare(sortValue(right));
    return byTime || number(left.callSeq) - number(right.callSeq);
  });
  return {
    calls: selected,
    rawCalls,
    rawCount: rawCalls.length,
    narrowed: selected.length < rawCalls.length,
  };
}

function chooseTiming(arrivalSelection, departureSelection) {
  const arrivals = arrivalSelection.calls.length ? arrivalSelection.calls : [null];
  const departures = departureSelection.calls.length ? departureSelection.calls : [null];
  const pairs = [];
  for (const arrivalCall of arrivals) {
    for (const departureCall of departures) {
      const start = dateMs(arrivalCall?.etb);
      const end = dateMs(departureCall?.etd);
      const hasBoth = Number.isFinite(start) && Number.isFinite(end);
      const gapHours = hasBoth ? (end - start) / 3600000 : null;
      const rank = hasBoth ? (gapHours >= 0 ? 0 : 1) : 2;
      pairs.push({
        arrivalCall,
        departureCall,
        gapHours,
        rank,
        score: hasBoth ? Math.abs(gapHours) : Number.POSITIVE_INFINITY,
      });
    }
  }
  pairs.sort((left, right) => left.rank - right.rank || left.score - right.score);
  return pairs[0] || { arrivalCall: null, departureCall: null, gapHours: null, rank: 2 };
}

function chooseEtbTiming(arrivalSelection, departureSelection) {
  const arrivals = arrivalSelection.calls.length ? arrivalSelection.calls : [null];
  const departures = departureSelection.calls.length ? departureSelection.calls : [null];
  const pairs = [];
  for (const arrivalCall of arrivals) {
    for (const departureCall of departures) {
      const start = dateMs(arrivalCall?.etb);
      const end = dateMs(departureCall?.etb);
      const hasBoth = Number.isFinite(start) && Number.isFinite(end);
      const gapHours = hasBoth ? (end - start) / 3600000 : null;
      const rank = hasBoth ? (gapHours >= 0 ? 0 : 1) : 2;
      pairs.push({
        arrivalCall,
        departureCall,
        gapHours,
        rank,
        score: hasBoth ? Math.abs(gapHours) : Number.POSITIVE_INFINITY,
      });
    }
  }
  pairs.sort((left, right) => left.rank - right.rank || left.score - right.score);
  return pairs[0] || { arrivalCall: null, departureCall: null, gapHours: null, rank: 2 };
}

function connectionRisk(timing) {
  if (!Number.isFinite(timing?.gapHours)) return "unknown";
  if (timing.gapHours < 0) return "missed";
  if (timing.gapHours < 24) return "tight";
  return "normal";
}

function transferType(pol, hub) {
  const polCountry = normalizedKey(pol).slice(0, 2);
  const hubCountry = normalizedKey(hub).slice(0, 2);
  if (polCountry.length < 2 || hubCountry.length < 2) return "unknown";
  return polCountry === hubCountry ? "domestic" : "international";
}

function callAtPortRelativeTo(anchorCall, direction, targetPort) {
  if (!anchorCall || !normalizedKey(targetPort)) return null;
  const { voyage, index } = voyageIndexForCall(anchorCall);
  if (index < 0) return null;
  const step = direction === "before" ? -1 : 1;
  for (let cursor = index + step; cursor >= 0 && cursor < voyage.length; cursor += step) {
    if (voyage[cursor].port === normalizedKey(targetPort)) return voyage[cursor];
  }
  return null;
}

function cargoSequence(arrivalLeg, departureLeg, hub, timing) {
  const previousPort = normalizedKey(arrivalLeg.pol) || "—";
  const nextPort = normalizedKey(departureLeg.pod) || "—";
  const previousCall =
    callAtPortRelativeTo(timing.arrivalCall, "before", previousPort) ||
    selectScheduleCalls(arrivalLeg, previousPort).calls[0] ||
    null;
  const nextCall =
    callAtPortRelativeTo(timing.departureCall, "after", nextPort) ||
    selectScheduleCalls(departureLeg, nextPort).calls[0] ||
    null;
  return {
    previous: { port: previousPort, call: previousCall },
    current: {
      port: normalizedKey(hub) || "—",
      arrivalCall: timing.arrivalCall,
      departureCall: timing.departureCall,
    },
    next: { port: nextPort, call: nextCall },
  };
}

const TRANSFER_SEARCH_FIELDS = {
  from: "Pre Port",
  to: "Post Port",
  bookingPol: "Booking POL",
  bookingPod: "Booking POD",
  laneIn: "LANE IN",
  laneOut: "LANE OUT",
  vvdIn: "VVD IN",
  vvdOut: "VVD OUT",
};

function scheduleCallSignature(call) {
  return [call?.lane, call?.vvd, call?.port, call?.eta, call?.etb, call?.etd, call?.callSeq, call?.sourceRow]
    .map((value) => String(value ?? ""))
    .join("|");
}

function uniqueScheduleCalls(calls) {
  const seen = new Set();
  return calls.filter((call) => {
    const signature = scheduleCallSignature(call);
    if (seen.has(signature)) return false;
    seen.add(signature);
    return true;
  }).sort((left, right) => {
    const leftTime = left.etb || left.eta || left.etd || "";
    const rightTime = right.etb || right.eta || right.etd || "";
    return leftTime.localeCompare(rightTime) || number(left.callSeq) - number(right.callSeq);
  });
}

function transferEventCallKey(event) {
  return joinKey(event.laneIn, event.vvdIn, event.port);
}

function indexTransferEvents() {
  state.transferEvents = [];
  state.transferEventsByPort.clear();
  for (const record of state.data.records) {
    for (let junctionIndex = 0; junctionIndex < record.legs.length - 1; junctionIndex += 1) {
      const arrivalLeg = record.legs[junctionIndex];
      const departureLeg = record.legs[junctionIndex + 1];
      const port = normalizedKey(arrivalLeg.pod || departureLeg.pol);
      if (!port) continue;
      const event = {
        key: `${record._id}|${junctionIndex}`,
        record,
        junctionIndex,
        port,
        from: normalizedKey(arrivalLeg.pol) || "—",
        to: normalizedKey(departureLeg.pod) || "—",
        bookingPol: normalizedKey(record.original?.pol) || normalizedKey(record.legs[0]?.pol) || "—",
        bookingPod: normalizedKey(record.original?.pod) || normalizedKey(record.legs.at(-1)?.pod) || "—",
        laneIn: normalizedKey(arrivalLeg.lane) || "—",
        laneOut: normalizedKey(departureLeg.lane) || "—",
        vvdIn: normalizedKey(arrivalLeg.vvd) || "—",
        vvdOut: normalizedKey(departureLeg.vvd) || "—",
        arrivalLeg,
        departureLeg,
        arrivalCalls: uniqueScheduleCalls(
          scheduleCallsFor(arrivalLeg, port).filter((call) => call.etb || call.eta),
        ),
      };
      event.callKey = transferEventCallKey(event);
      state.transferEvents.push(event);
      if (!state.transferEventsByPort.has(port)) state.transferEventsByPort.set(port, []);
      state.transferEventsByPort.get(port).push(event);
    }
  }
  const ports = [...state.transferEventsByPort.entries()]
    .sort((left, right) => right[1].length - left[1].length || compareAlpha(left[0], right[0]))
    .map(([port]) => port);
  els.transferPortList.innerHTML = optionList(ports);
}

function transferSearchFilterValue(group) {
  return group ? filterValue(group) : "ALL";
}

function eventPassesTransferCargoFilters(event) {
  const coc = transferSearchFilterValue(els.transferCocFilter);
  const sul = transferSearchFilterValue(els.transferSulFilter);
  if (coc !== "ALL" && normalizedKey(event.record.coc) !== coc) return false;
  if (sul !== "ALL" && normalizedKey(event.record.sul) !== sul) return false;
  return true;
}

function eventPassesTransferOptionalFilters(event, ignoredField = "") {
  for (const field of Object.keys(TRANSFER_SEARCH_FIELDS)) {
    if (field === ignoredField) continue;
    const selected = state.transferSearchFilters[field];
    if (selected?.size && !selected.has(normalizedKey(event[field]))) return false;
  }
  return true;
}

function transferDateValue(call) {
  return String(call?.etb || "").slice(0, 10);
}

function transferDateInRange(value, start, end) {
  return Boolean(value && (!start || value >= start) && (!end || value <= end));
}

function normalizeTransferDateInput(input) {
  const raw = input.value.trim();
  if (!raw) return "";
  const digits = raw.replace(/\D/g, "");
  if (digits.length !== 8) return raw;
  const normalized = `${digits.slice(0, 4)}-${digits.slice(4, 6)}-${digits.slice(6, 8)}`;
  input.value = normalized;
  return normalized;
}

function validTransferDate(value) {
  if (!value) return true;
  const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value);
  if (!match) return false;
  const date = new Date(`${value}T00:00:00`);
  return !Number.isNaN(date.getTime())
    && date.getFullYear() === Number(match[1])
    && date.getMonth() + 1 === Number(match[2])
    && date.getDate() === Number(match[3]);
}

function hasTransferRouteFilter() {
  return Object.values(state.transferSearchFilters).some((values) => values.size > 0);
}

function hasTransferSearchScope(start, end) {
  return Boolean(start || end || hasTransferRouteFilter());
}

function markTransferSearchDirty() {
  state.transferSearchSubmitted = false;
}

function selectedArrivalCall(event) {
  if (!event.arrivalCalls.length) return null;
  if (event.arrivalCalls.length === 1) return event.arrivalCalls[0];
  const selected = state.transferCallSelections.get(event.callKey);
  return event.arrivalCalls.find((call) => scheduleCallSignature(call) === selected) || null;
}

function transferSearchCandidateEvents(port) {
  return (state.transferEventsByPort.get(port) || []).filter(eventPassesTransferCargoFilters);
}

function buildTransferSearchView() {
  const port = normalizedKey(els.transferPortInput.value);
  const start = els.transferStartDate.value.trim();
  const end = els.transferEndDate.value.trim();
  const resolved = [];
  const missing = [];
  const pendingByCall = new Map();
  if (!port || !hasTransferSearchScope(start, end) || !validTransferDate(start) || !validTransferDate(end) || (start && end && start > end)) {
    return { port, start, end, resolved, missing, pending: [] };
  }

  for (const event of transferSearchCandidateEvents(port)) {
    if (!eventPassesTransferOptionalFilters(event)) continue;
    if (!event.arrivalCalls.length) {
      missing.push(event);
      continue;
    }
    const selectedCall = selectedArrivalCall(event);
    if (event.arrivalCalls.length > 1 && !selectedCall) {
      if (!event.arrivalCalls.some((call) => transferDateInRange(transferDateValue(call), start, end))) continue;
      if (!pendingByCall.has(event.callKey)) {
        pendingByCall.set(event.callKey, { key: event.callKey, event, events: [], calls: event.arrivalCalls });
      }
      pendingByCall.get(event.callKey).events.push(event);
      continue;
    }
    const arrivalDate = transferDateValue(selectedCall);
    if (!transferDateInRange(arrivalDate, start, end)) continue;
    const timing = chooseTiming(
      { calls: [selectedCall] },
      selectScheduleCalls(event.departureLeg, event.port),
    );
    resolved.push({ ...event, selectedCall, arrivalDate, timing, risk: connectionRisk(timing) });
  }
  return {
    port,
    start,
    end,
    resolved,
    missing,
    pending: [...pendingByCall.values()].sort((left, right) => compareAlpha(left.event.vvdIn, right.event.vvdIn)),
  };
}

function totalsForTransferEvents(events) {
  const totals = newTotals();
  for (const event of events) addRecordTotals(totals, event.record);
  return totals;
}

function transferSearchFilterOptions(field) {
  const port = normalizedKey(els.transferPortInput.value);
  if (!port) return [];
  const values = new Set();
  for (const event of transferSearchCandidateEvents(port)) {
    if (!eventPassesTransferOptionalFilters(event, field)) continue;
    const value = normalizedKey(event[field]);
    if (value && value !== "—") values.add(value);
  }
  return [...values].sort(compareAlpha);
}

function compareScheduleValues(leftValue, rightValue, fallbackLeft, fallbackRight) {
  if (leftValue && rightValue && leftValue !== rightValue) return leftValue.localeCompare(rightValue);
  if (leftValue && !rightValue) return -1;
  if (!leftValue && rightValue) return 1;
  return compareAlpha(fallbackLeft, fallbackRight);
}

function buildDirectGroups(matches) {
  const groups = new Map();
  for (const { record, legIndex } of matches) {
    if (record.legs.length !== 1) continue;
    const leg = record.legs[legIndex];
    const pol = normalizedKey(leg.pol) || "—";
    const pod = normalizedKey(leg.pod) || "—";
    const polKey = pol;
    if (!groups.has(polKey)) {
      groups.set(polKey, {
        key: polKey,
        pol,
        totals: newTotals(),
        recordIds: new Set(),
        etds: new Set(),
        pods: new Map(),
      });
    }
    const group = groups.get(polKey);
    if (!group.pods.has(pod)) {
      group.pods.set(pod, {
        key: `${polKey}|${pod}`,
        pod,
        totals: newTotals(),
        recordIds: new Set(),
        etds: new Set(),
        candidateCount: 0,
      });
    }
    const podGroup = group.pods.get(pod);
    const selection = selectScheduleCalls(leg, leg.pol);
    selection.calls.forEach((call) => {
      if (call.etd) {
        group.etds.add(call.etd);
        podGroup.etds.add(call.etd);
      }
    });
    podGroup.candidateCount = Math.max(podGroup.candidateCount, selection.rawCount);
    addRecordOnce(group, record);
    addRecordOnce(podGroup, record);
  }

  return [...groups.values()]
    .map((group) => ({
      ...group,
      pods: [...group.pods.values()].sort((left, right) =>
        compareScheduleValues(earliest(left.etds), earliest(right.etds), left.pod, right.pod),
      ),
    }))
    .sort((left, right) =>
      compareScheduleValues(earliest(left.etds), earliest(right.etds), left.pol, right.pol),
    );
}

function routeSignature(leg) {
  return joinKey(leg.pol, leg.pod, leg.lane, leg.vvd);
}

function buildConnectionDescriptor(record, legIndex, kind) {
  const focusLeg = record.legs[legIndex];
  const isInbound = kind === "inbound";
  const arrivalLeg = isInbound ? record.legs[legIndex - 1] : focusLeg;
  const departureLeg = isInbound ? focusLeg : record.legs[legIndex + 1];
  const hub = normalizedKey(isInbound ? focusLeg.pol : focusLeg.pod);
  if (!arrivalLeg || !departureLeg || !hub) return null;

  const arrivalSelection = selectScheduleCalls(arrivalLeg, hub);
  const departureSelection = selectScheduleCalls(departureLeg, hub);
  const timing = chooseEtbTiming(arrivalSelection, departureSelection);
  const sequence = cargoSequence(arrivalLeg, departureLeg, hub, timing);
  const key = [kind, hub, routeSignature(arrivalLeg), routeSignature(departureLeg)].join("|");
  const segmentStart = normalizedKey(focusLeg.pol) || "—";
  const segmentEnd = normalizedKey(focusLeg.pod) || "—";
  const segmentKey = [kind, segmentStart, segmentEnd, normalizedKey(focusLeg.lane), normalizedKey(focusLeg.vvd)].join("|");

  return {
    key,
    kind,
    hub,
    focusLeg,
    segmentStart,
    segmentEnd,
    segmentKey,
    arrivalLeg,
    departureLeg,
    arrivalSelection,
    departureSelection,
    timing,
    risk: connectionRisk(timing),
    transferType: transferType(arrivalLeg.pol, hub),
    sequence,
  };
}

function addOriginalOd(connection, record) {
  const origin = normalizedKey(record.legs[0]?.pol) || "—";
  const destination = normalizedKey(record.legs.at(-1)?.pod) || "—";
  const key = joinKey(origin, destination);
  if (!connection.originalOds.has(key)) {
    connection.originalOds.set(key, {
      key,
      origin,
      destination,
      totals: newTotals(),
      recordIds: new Set(),
    });
  }
  addRecordOnce(connection.originalOds.get(key), record);
}

function matchSupportsTransferDirection(match, direction) {
  if (direction === "inbound") return match.legIndex > 0;
  if (direction === "outbound") return match.legIndex < match.record.legs.length - 1;
  return match.record.legs.length > 1;
}

function filterTransferMatches(matches, direction = state.transferDirection) {
  return matches.filter((match) => matchSupportsTransferDirection(match, direction));
}

function buildTransferGroups(matches, direction = state.transferDirection) {
  const groups = new Map();
  for (const { record, legIndex } of matches) {
    if (record.legs.length <= 1) continue;
    const descriptors = [];
    if (direction !== "outbound" && legIndex > 0) {
      descriptors.push(buildConnectionDescriptor(record, legIndex, "inbound"));
    }
    if (direction !== "inbound" && legIndex < record.legs.length - 1) {
      descriptors.push(buildConnectionDescriptor(record, legIndex, "outbound"));
    }

    for (const descriptor of descriptors.filter(Boolean)) {
      if (!groups.has(descriptor.hub)) {
        groups.set(descriptor.hub, {
          port: descriptor.hub,
          totals: newTotals(),
          recordIds: new Set(),
          connections: new Map(),
          segments: new Map(),
        });
      }
      const group = groups.get(descriptor.hub);
      addRecordOnce(group, record);
      if (!group.connections.has(descriptor.key)) {
        group.connections.set(descriptor.key, {
          ...descriptor,
          totals: newTotals(),
          recordIds: new Set(),
          originalOds: new Map(),
          records: new Map(),
        });
      }
      const connection = group.connections.get(descriptor.key);
      addRecordOnce(connection, record);
      connection.records.set(record._id, record);
      addOriginalOd(connection, record);
      if (!group.segments.has(descriptor.segmentKey)) {
        group.segments.set(descriptor.segmentKey, {
          key: descriptor.segmentKey,
          kind: descriptor.kind,
          start: descriptor.segmentStart,
          end: descriptor.segmentEnd,
          lane: normalizedKey(descriptor.focusLeg.lane),
          vvd: normalizedKey(descriptor.focusLeg.vvd),
          totals: newTotals(),
          recordIds: new Set(),
          connectionKeys: new Set(),
        });
      }
      const segment = group.segments.get(descriptor.segmentKey);
      addRecordOnce(segment, record);
      segment.connectionKeys.add(descriptor.key);
    }
  }

  return [...groups.values()]
    .map((group) => {
      const connections = [...group.connections.values()]
        .map((connection) => ({
          ...connection,
          records: [...connection.records.values()],
          originalOds: [...connection.originalOds.values()].sort((left, right) =>
            compareAlpha(`${left.origin}|${left.destination}`, `${right.origin}|${right.destination}`),
          ),
        }))
        .sort((left, right) => {
          if (left.kind !== right.kind) return left.kind === "inbound" ? -1 : 1;
          return compareScheduleValues(
            left.timing.departureCall?.etb || "",
            right.timing.departureCall?.etb || "",
            left.key,
            right.key,
          );
        });
      const connectionByKey = new Map(connections.map((connection) => [connection.key, connection]));
      const segments = [...group.segments.values()]
        .map((segment) => ({
          ...segment,
          connections: [...segment.connectionKeys]
            .map((key) => connectionByKey.get(key))
            .filter(Boolean),
        }))
        .sort((left, right) => compareAlpha(`${left.start}|${left.end}|${left.kind}`, `${right.start}|${right.end}|${right.kind}`));
      return { ...group, connections, segments };
    })
    .sort((left, right) => compareAlpha(left.port, right.port));
}

function metricTriplet(totals, prefix) {
  return `
    <span class="metric-triplet" aria-label="20FT、40FT、TEU">
      <span><small>20</small><strong>${display(totals[`${prefix}20`])}</strong></span>
      <span><small>40</small><strong>${display(totals[`${prefix}40`])}</strong></span>
      <span class="metric-teu teu"><small>TEU</small><strong>${display(totals[`${prefix}Teu`])}</strong></span>
    </span>
  `;
}

function directMetricTriplet(totals, prefix) {
  return `
    <span class="metric-triplet direct-number-triplet" aria-label="20FT ${display(totals[`${prefix}20`])}、40FT ${display(totals[`${prefix}40`])}、TEU ${display(totals[`${prefix}Teu`])}">
      <strong>${display(totals[`${prefix}20`])}</strong>
      <strong>${display(totals[`${prefix}40`])}</strong>
      <strong class="metric-teu teu">${display(totals[`${prefix}Teu`])}</strong>
    </span>
  `;
}

function metricTeu(totals, prefix) {
  return `
    <span class="metric-single metric-teu" aria-label="TEU ${display(totals[`${prefix}Teu`])}">
      <strong>${display(totals[`${prefix}Teu`])}</strong>
    </span>
  `;
}

function cargoMatrix(totals, label) {
  return `
    <div class="cargo-summary" aria-label="${escapeHtml(label)}">
      <div class="cargo-matrix">
        <span class="cargo-corner cargo-head cargo-label">箱型</span><strong class="cargo-head">Booking</strong><strong class="cargo-head">OP</strong><strong class="cargo-head">VL</strong>
        <span class="cargo-label">20FT</span><strong>${display(totals.b20)}</strong><strong>${display(totals.op20)}</strong><strong>${display(totals.vl20)}</strong>
        <span class="cargo-label">40FT</span><strong>${display(totals.b40)}</strong><strong>${display(totals.op40)}</strong><strong>${display(totals.vl40)}</strong>
        <span class="cargo-label cargo-teu-label">TEU</span><strong class="cargo-teu">${display(totals.bTeu)}</strong><strong class="cargo-teu">${display(totals.opTeu)}</strong><strong class="cargo-teu">${display(totals.vlTeu)}</strong>
      </div>
      <div class="cargo-weight weight-tile"><span>Weight TON</span><strong>${display(totals.weight)}</strong></div>
    </div>
  `;
}

function etdCell(values) {
  const sorted = [...values].filter(Boolean).sort();
  if (!sorted.length) return `<span class="schedule-missing">船期暂无</span>`;
  return `<span class="etd-value etd-cell"><strong>${escapeHtml(formatDateOnly(sorted[0], true))}</strong></span>`;
}

function renderDirect(groups) {
  els.directBadge.textContent = `${display(groups.length, true)} POL`;
  if (!groups.length) {
    els.directRows.innerHTML = `
      <div class="module-empty">
        <span class="module-empty-mark">D</span>
        <div><strong>当前筛选没有直航记录</strong><p>可切换到“中转监控”查看多程 Booking。</p></div>
      </div>`;
    return;
  }

  const visibleGroups = groups.slice(0, 200);
  els.directRows.innerHTML = visibleGroups
    .map((group) => {
      const expanded = state.expandedPols.has(group.key);
      return `
        <article class="direct-group${expanded ? " expanded" : ""}">
          <button class="direct-group-header direct-group-summary" type="button" data-pol-toggle="${escapeHtml(group.key)}" aria-expanded="${expanded}">
            <span class="direct-route-cell direct-port-cell"><span class="expand-indicator direct-expand" aria-hidden="true">${expanded ? "−" : "+"}</span><span><strong>${escapeHtml(group.pol)}</strong><small>${display(group.pods.length, true)} POD</small></span></span>
            ${etdCell(group.etds)}
            ${directMetricTriplet(group.totals, "b")}
            ${metricTeu(group.totals, "op")}
            ${metricTeu(group.totals, "vl")}
            <span class="direct-weight weight-cell"><strong>${display(group.totals.weight)}</strong><small>TON</small></span>
          </button>
          <div class="direct-group-body"${expanded ? "" : " hidden"}>
            ${group.pods
              .map(
                (pod) => `
                  <div class="pod-row">
                    <span class="direct-route-cell direct-port-cell pod-cell"><span class="pod-branch" aria-hidden="true">↳</span><span><strong>${escapeHtml(pod.pod)}</strong></span></span>
                    ${etdCell(pod.etds)}
                    ${directMetricTriplet(pod.totals, "b")}
                    ${metricTeu(pod.totals, "op")}
                    ${metricTeu(pod.totals, "vl")}
                    <span class="direct-weight weight-cell"><strong>${display(pod.totals.weight)}</strong><small>TON</small></span>
                  </div>`,
              )
              .join("")}
          </div>
        </article>`;
    })
    .join("");

  if (groups.length > visibleGroups.length) {
    els.directRows.insertAdjacentHTML(
      "beforeend",
      `<p class="limit-note">当前显示前 ${display(visibleGroups.length, true)} 个 POL，请继续收窄筛选。</p>`,
    );
  }

  const totals = groups.reduce((sum, group) => {
    addTotals(sum, group.totals);
    return sum;
  }, newTotals());
  els.directRows.insertAdjacentHTML(
    "beforeend",
    `<div class="direct-total-row" role="row" aria-label="直航合计">
      <span class="direct-total-label"><strong>TOTAL</strong></span>
      <span aria-hidden="true"></span>
      ${directMetricTriplet(totals, "b")}
      ${metricTeu(totals, "op")}
      ${metricTeu(totals, "vl")}
      <span class="direct-weight weight-cell"><strong>${display(totals.weight)}</strong><small>TON</small></span>
    </div>`,
  );
}

function ensureSelectedTransferPort(groups) {
  if (!groups.length) {
    state.selectedTransferPort = "";
    return null;
  }
  const selected = groups.find((group) => group.port === state.selectedTransferPort);
  if (selected) return selected;
  state.selectedTransferPort = groups[0].port;
  return groups[0];
}

const TRANSFER_RISK_META = {
  all: { label: "全部", short: "全部", className: "all" },
  missed: { label: "接不上", short: "接不上", className: "missed" },
  tight: { label: "小于 24h", short: "<24h", className: "tight" },
  normal: { label: "正常", short: "正常", className: "normal" },
  unknown: { label: "船期待确认", short: "待确认", className: "unknown" },
};

const TRANSFER_TYPE_META = {
  all: { label: "全部类型", short: "全部" },
  domestic: { label: "Domestic T/S", short: "Domestic" },
  international: { label: "International T/S", short: "International" },
  unknown: { label: "待确认", short: "待确认" },
};

function shortPort(port) {
  const value = normalizedKey(port);
  return value.length === 5 ? value.slice(2) : value || "—";
}

function domId(prefix, value) {
  let hash = 2166136261;
  for (const character of String(value || "")) {
    hash ^= character.charCodeAt(0);
    hash = Math.imul(hash, 16777619);
  }
  return `${prefix}-${(hash >>> 0).toString(36)}`;
}

function transferMetricGrid(totals, className = "") {
  return `
    <span class="transfer-metric-grid transfer-primary-metrics ${className}" aria-label="Booking TEU 与含皮重量">
      <span class="metric-primary-teu"><small>TEU</small><strong>${display(totals.bTeu)}</strong></span>
      <span class="metric-primary-weight"><small>TON</small><strong>${display(totals.weight)}</strong></span>
    </span>`;
}

function transferFullMetricGrid(totals, className = "") {
  return `
    <span class="transfer-metric-grid transfer-full-metric-grid ${className}" aria-label="Booking、OP、VL 与含皮重量">
      <span><small>20'</small><strong>${display(totals.b20)}</strong></span>
      <span><small>40'</small><strong>${display(totals.b40)}</strong></span>
      <span class="metric-bkg-teu"><small>BKG TEU</small><strong>${display(totals.bTeu)}</strong></span>
      <span class="metric-op-teu"><small>OP TEU</small><strong>${display(totals.opTeu)}</strong></span>
      <span class="metric-vl-teu"><small>VL TEU</small><strong>${display(totals.vlTeu)}</strong></span>
      <span class="metric-primary-weight"><small>TON</small><strong>${display(totals.weight)}</strong></span>
    </span>`;
}

function totalsForConnections(connections) {
  const totals = newTotals();
  const recordIds = new Set();
  for (const connection of connections) {
    for (const record of connection.records) {
      if (recordIds.has(record._id)) continue;
      recordIds.add(record._id);
      addRecordTotals(totals, record);
    }
  }
  return { totals, count: recordIds.size };
}

function connectionsForRisk(connections, risk = state.transferRisk) {
  return risk === "all" ? connections : connections.filter((connection) => connection.risk === risk);
}

function connectionsForType(connections, type = state.transferType) {
  return type === "all" ? connections : connections.filter((connection) => connection.transferType === type);
}

function connectionsForTransferFilters(
  connections,
  risk = state.transferRisk,
  type = state.transferType,
) {
  return connectionsForType(connectionsForRisk(connections, risk), type);
}

function riskSummary(connections) {
  return Object.keys(TRANSFER_RISK_META).reduce((summary, risk) => {
    const selected = connectionsForTransferFilters(connections, risk, state.transferType);
    summary[risk] = { connections: selected.length, ...totalsForConnections(selected) };
    return summary;
  }, {});
}

function typeSummary(connections) {
  return Object.keys(TRANSFER_TYPE_META).reduce((summary, type) => {
    const selected = connectionsForTransferFilters(connections, state.transferRisk, type);
    summary[type] = { connections: selected.length, ...totalsForConnections(selected) };
    return summary;
  }, {});
}

function segmentView(segment, risk = state.transferRisk, type = state.transferType) {
  const connections = connectionsForTransferFilters(segment.connections, risk, type);
  if (!connections.length) return null;
  return { ...segment, connections, ...totalsForConnections(connections) };
}

function renderTransferPortTabs(groups) {
  els.transferBadge.textContent = `${display(groups.length, true)} T/S 港`;
  if (!groups.length) {
    els.transferPortTabs.innerHTML = "";
    return;
  }
  const selected = ensureSelectedTransferPort(groups);
  els.transferPortTabs.innerHTML = groups
    .map((group) => {
      const active = group.port === selected?.port;
      return `
        <button class="transfer-port-button${active ? " active" : ""}" type="button" data-ts-port="${escapeHtml(group.port)}" aria-pressed="${active}">
          <span class="transfer-port-identity"><small>T/S PORT</small><strong>${escapeHtml(group.port)}</strong></span>
          ${transferMetricGrid(group.totals, "port-tab-metrics")}
        </button>`;
    })
    .join("");
}

function riskBadge(risk, totals = null) {
  const meta = TRANSFER_RISK_META[risk] || TRANSFER_RISK_META.unknown;
  return `<span class="transfer-risk-badge ${meta.className}"><strong>${meta.short}</strong>${totals ? `<small>${display(totals.bTeu)} TEU</small>` : ""}</span>`;
}

function transferTypeBadge(type) {
  const label = type === "domestic" ? "Domestic T/S" : type === "international" ? "International T/S" : "T/S 待确认";
  return `<span class="transfer-type-badge ${type}">${label}</span>`;
}

function formatEtbGap(hours) {
  if (!Number.isFinite(hours)) return "待确认";
  const rounded = Math.round(hours);
  if (rounded < 0) return `晚到 ${display(Math.abs(rounded))}h`;
  if (rounded < 24) return `${display(rounded)}h`;
  const days = Math.floor(rounded / 24);
  const remaining = rounded % 24;
  return remaining ? `${days}d ${remaining}h` : `${days}d`;
}

function recordPath(record) {
  const ports = [];
  for (const leg of record.legs) {
    const pol = normalizedKey(leg.pol) || "—";
    const pod = normalizedKey(leg.pod) || "—";
    if (!ports.length || ports.at(-1) !== pol) ports.push(pol);
    ports.push(pod);
  }
  return ports.join("-");
}

function blRows(connection) {
  const records = connection.records.slice(0, 80);
  return `
    <div class="transfer-bl-wrap">
      <div class="transfer-bl-head"><span>BL No.</span><span>CUL CODE</span><span>20'</span><span>40'</span><span>TEU</span><span>OP</span><span>VL</span><span>TON</span><span>完整路径</span></div>
      ${records.map((record) => `
        <div class="transfer-bl-row">
          <strong>${escapeHtml(record.blNo || "—")}</strong>
          <span>${escapeHtml(record.culCode || "—")}</span>
          <span>${display(record.b.t20)}</span><span>${display(record.b.t40)}</span><span>${display(record.b.teu)}</span>
          <span>${display(record.op.teu)}</span><span>${display(record.vl.teu)}</span><span>${display(record.w)}</span>
          <span class="transfer-bl-path">${escapeHtml(recordPath(record))}</span>
        </div>`).join("")}
      ${connection.records.length > records.length ? `<p class="transfer-bl-limit">网页显示前 ${display(records.length, true)} 条，完整 ${display(connection.records.length, true)} 条请使用 Excel 导出。</p>` : ""}
    </div>`;
}

function feederConnectionRow(connection) {
  const expanded = state.expandedConnections.has(connection.key);
  const inbound = connection.kind === "inbound";
  const firstLabel = inbound ? "前程船" : "当前 VVD";
  const secondLabel = inbound ? "当前 VVD" : "后程船";
  const detailsId = domId("connection-bl", connection.key);
  return `
    <article class="feeder-connection ${connection.risk}">
      <button type="button" class="feeder-connection-summary" data-connection-toggle="${escapeHtml(connection.key)}" aria-expanded="${expanded}" aria-controls="${detailsId}">
        <span class="feeder-status">${riskBadge(connection.risk)}${transferTypeBadge(connection.transferType)}</span>
        <span class="feeder-voyage"><small>${firstLabel}</small><strong>${escapeHtml(connection.arrivalLeg.vvd || "—")}</strong><em>${escapeHtml(connection.arrivalLeg.lane || "—")}</em></span>
        <span class="feeder-time"><small>${firstLabel} ETB</small><strong>${escapeHtml(formatDate(connection.timing.arrivalCall?.etb, true))}</strong></span>
        <span class="feeder-time"><small>${secondLabel} ETB</small><strong>${escapeHtml(formatDate(connection.timing.departureCall?.etb, true))}</strong></span>
        <span class="feeder-gap"><small>接驳</small><strong>${escapeHtml(formatEtbGap(connection.timing.gapHours))}</strong></span>
        ${transferFullMetricGrid(connection.totals, "connection-full-metrics")}
        <span class="feeder-expand">${expanded ? "收起 BL" : "查看 BL"}</span>
      </button>
      <div id="${detailsId}" class="feeder-bl-details"${expanded ? "" : " hidden"}>${blRows(connection)}</div>
    </article>`;
}

function feederGroups(segment) {
  const groups = new Map();
  for (const connection of segment.connections) {
    const pol = normalizedKey(connection.arrivalLeg.pol) || "—";
    if (!groups.has(pol)) groups.set(pol, { pol, connections: [] });
    groups.get(pol).connections.push(connection);
  }
  return [...groups.values()]
    .map((group) => ({ ...group, ...totalsForConnections(group.connections) }))
    .sort((left, right) => right.totals.bTeu - left.totals.bTeu || compareAlpha(left.pol, right.pol));
}

function transferDrawer(group, segment) {
  if (!segment) return "";
  const directionLabel = segment.kind === "inbound" ? "转入当前 VVD" : "转出当前 VVD";
  const groups = feederGroups(segment);
  if (state.drawerSegmentKey !== segment.key) {
    state.drawerSegmentKey = segment.key;
    state.expandedFeederGroups.clear();
    state.expandedConnections.clear();
    if (groups[0]) state.expandedFeederGroups.add(groups[0].pol);
  }
  return `
    <div class="transfer-drawer-backdrop transfer-modal-backdrop" data-transfer-drawer-close></div>
    <aside class="transfer-drawer transfer-modal" role="dialog" aria-modal="true" aria-labelledby="transferDrawerTitle">
      <header class="transfer-drawer-head">
        <div>
          <span>${directionLabel}</span>
          <h3 id="transferDrawerTitle">${escapeHtml(shortPort(segment.start))} → ${escapeHtml(shortPort(segment.end))}</h3>
          <p>${escapeHtml(segment.start)} → ${escapeHtml(segment.end)}　${display(segment.count, true)} 条 BL</p>
        </div>
        <button type="button" class="drawer-close" data-transfer-drawer-close aria-label="关闭接驳弹窗">×</button>
      </header>
      <div class="transfer-drawer-metrics">
        ${transferFullMetricGrid(segment.totals, "drawer-full-metrics")}
      </div>
      <div class="transfer-drawer-toolbar">
        <p>先展开前程 POL，再按需查看 BL；展开后当前位置保持不变。接驳只比较双方在 ${escapeHtml(group.port)} 的 ETB。</p>
        <button type="button" data-export-scope="segment">导出此航段</button>
      </div>
      <div class="transfer-drawer-body">
        ${groups.map((feederGroup) => {
          const expanded = state.expandedFeederGroups.has(feederGroup.pol);
          const groupId = domId("feeder-pol", `${segment.key}|${feederGroup.pol}`);
          return `
          <section class="feeder-pol-group${expanded ? " expanded" : ""}">
            <button type="button" class="feeder-pol-toggle" data-feeder-group-toggle="${escapeHtml(feederGroup.pol)}" aria-expanded="${expanded}" aria-controls="${groupId}">
              <span class="feeder-pol-identity"><span class="feeder-pol-chevron" aria-hidden="true">${expanded ? "−" : "+"}</span><span><small>前程 POL</small><strong>${escapeHtml(feederGroup.pol)}</strong><em>${display(feederGroup.connections.length, true)} 个前程连接</em></span></span>
              ${transferMetricGrid(feederGroup.totals, "feeder-pol-metrics")}
            </button>
            <div id="${groupId}" class="feeder-connections"${expanded ? "" : " hidden"}>${feederGroup.connections.map(feederConnectionRow).join("")}</div>
          </section>`;
        }).join("")}
      </div>
    </aside>`;
}

function segmentRiskBadges(segment) {
  const summary = riskSummary(segment.connections);
  return ["missed", "tight", "unknown"]
    .filter((risk) => summary[risk].connections)
    .map((risk) => riskBadge(risk, summary[risk].totals))
    .join("") || '<span class="transfer-risk-clear">无异常</span>';
}

function transferSegmentRow(segment) {
  const directionLabel = segment.kind === "inbound" ? "转入" : "转出";
  return `
    <button type="button" class="transfer-segment-row" data-transfer-segment="${escapeHtml(segment.key)}">
      <span class="segment-route"><small>${directionLabel}当前 VVD</small><strong>${escapeHtml(shortPort(segment.start))} → ${escapeHtml(shortPort(segment.end))}</strong><em>${escapeHtml(segment.start)} → ${escapeHtml(segment.end)}</em></span>
      <span class="segment-bkg-teu"><small>BKG TEU</small><strong>${display(segment.totals.bTeu)}</strong></span>
      <span class="segment-op-teu"><small>OP TEU</small><strong>${display(segment.totals.opTeu)}</strong></span>
      <span class="segment-vl-teu"><small>VL TEU</small><strong>${display(segment.totals.vlTeu)}</strong></span>
      <span class="segment-weight"><small>Weight TON</small><strong>${display(segment.totals.weight)}</strong></span>
      <span class="segment-risks">${segmentRiskBadges(segment)}</span>
      <span class="segment-open">查看</span>
    </button>`;
}

function transferRiskControls(group) {
  const summary = riskSummary(group.connections);
  return Object.entries(TRANSFER_RISK_META).map(([risk, meta]) => {
    const active = risk === state.transferRisk;
    return `<button type="button" class="transfer-risk-option ${meta.className}${active ? " active" : ""}" data-transfer-risk="${risk}" aria-pressed="${active}"><strong>${meta.label}</strong><small>${display(summary[risk].connections, true)} 个前程</small></button>`;
  }).join("");
}

function transferTypeControls(group) {
  const summary = typeSummary(group.connections);
  return Object.entries(TRANSFER_TYPE_META).map(([type, meta]) => {
    const active = type === state.transferType;
    return `<button type="button" class="transfer-type-option ${type}${active ? " active" : ""}" data-transfer-type="${type}" aria-pressed="${active}"><strong>${meta.label}</strong><small>${display(summary[type].connections, true)} 个前程</small></button>`;
  }).join("");
}

function renderTransfer(groups) {
  const group = ensureSelectedTransferPort(groups);
  if (!group) {
    document.body.classList.remove("transfer-modal-open");
    els.transferRows.innerHTML = `
      <div class="module-empty">
        <span class="module-empty-mark">T/S</span>
        <div><strong>当前筛选没有中转记录</strong><p>可切换到“直航监控”查看单程 Booking。</p></div>
      </div>`;
    return;
  }

  const segments = group.segments.map((segment) => segmentView(segment)).filter(Boolean);
  const selectedSegment = segments.find((segment) => segment.key === state.selectedTransferSegment) || null;
  if (state.selectedTransferSegment && !selectedSegment) state.selectedTransferSegment = "";
  document.body.classList.toggle("transfer-modal-open", Boolean(selectedSegment));
  const filtered = totalsForConnections(connectionsForTransferFilters(group.connections));
  const activeFilters = [];
  if (state.transferRisk !== "all") activeFilters.push(TRANSFER_RISK_META[state.transferRisk].label);
  if (state.transferType !== "all") activeFilters.push(TRANSFER_TYPE_META[state.transferType].label);
  const filteredCaption = activeFilters.length
    ? `当前筛选　${activeFilters.join(" / ")}　${display(filtered.count, true)} 条 BL`
    : "显示全部接驳货量";
  els.transferRows.innerHTML = `
    <div class="transfer-port-overview">
      <div><span>SELECTED T/S PORT</span><h3>${escapeHtml(group.port)} 中转</h3><p>${display(group.recordIds.size, true)} 条唯一 BL　总体货量不随风险筛选变化</p></div>
      ${transferMetricGrid(group.totals, "port-overview-metrics")}
    </div>
    <section class="transfer-risk-workbench" aria-label="中转接驳筛选">
      <div class="risk-workbench-copy"><span>接驳筛选</span><p>可同时选择接驳状态与 Domestic / International T/S</p></div>
      <div class="transfer-filter-stack">
        <div class="transfer-filter-row"><span>状态</span><div class="transfer-risk-options">${transferRiskControls(group)}</div></div>
        <div class="transfer-filter-row"><span>类型</span><div class="transfer-type-options">${transferTypeControls(group)}</div></div>
      </div>
      <div class="transfer-export-actions"><button type="button" data-export-scope="current">导出当前视图</button><button type="button" data-export-scope="all">导出全部中转</button></div>
      <p class="transfer-export-status" role="status" aria-live="polite"></p>
    </section>
    <section class="transfer-segment-panel">
      <header><div><span>CURRENT VVD SEGMENTS</span><h3>当前 VVD 航段汇总</h3><p>${filteredCaption}</p></div>${transferMetricGrid(filtered.totals, "filtered-view-metrics")}</header>
      <div class="transfer-segment-head" aria-hidden="true"><span>当前 VVD 航段</span><span>BKG TEU</span><span>OP TEU</span><span>VL TEU</span><span>Weight TON</span><span>风险货量</span><span></span></div>
      <div class="transfer-segment-list">${segments.length ? segments.map(transferSegmentRow).join("") : `<div class="transfer-segment-empty"><strong>当前条件没有对应接驳</strong><p>可切换接驳状态或中转港查看其他结果。</p></div>`}</div>
    </section>
    ${transferDrawer(group, selectedSegment)}`;
}

function transferDirectionLabel(kind) {
  return kind === "inbound" ? "转入当前 VVD" : "转出当前 VVD";
}

function transferRiskLabel(risk) {
  return TRANSFER_RISK_META[risk]?.label || TRANSFER_RISK_META.unknown.label;
}

function transferTypeLabel(type) {
  if (type === "domestic") return "Domestic T/S";
  if (type === "international") return "International T/S";
  return "T/S 待确认";
}

function exportTransferSelection(scope) {
  const transferMatches = collectMatches().filter(({ record }) => record.legs.length > 1);
  const groups = buildTransferGroups(transferMatches, scope === "all" ? "all" : state.transferDirection);
  const selectedPort = scope === "all" ? "" : state.selectedTransferPort;
  const risk = scope === "all" ? "all" : state.transferRisk;
  const type = scope === "all" ? "all" : state.transferType;
  const selectedSegment = scope === "segment" ? state.selectedTransferSegment : "";
  const selections = [];
  for (const group of groups) {
    if (selectedPort && group.port !== selectedPort) continue;
    for (const segment of group.segments) {
      if (selectedSegment && segment.key !== selectedSegment) continue;
      const view = segmentView(segment, risk, type);
      if (view) selections.push({ group, segment: view });
    }
  }
  return selections;
}

function excelNumber(value) {
  return Math.round(number(value) * 10) / 10;
}

function setWorksheetWidths(XLSX, worksheet, widths) {
  worksheet["!cols"] = widths.map((wch) => ({ wch }));
  if (worksheet["!ref"]) worksheet["!autofilter"] = { ref: worksheet["!ref"] };
  if (!worksheet["!ref"]) return;
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  const headers = [];
  for (let column = range.s.c; column <= range.e.c; column += 1) {
    headers[column] = worksheet[XLSX.utils.encode_cell({ r: range.s.r, c: column })]?.v || "";
  }
  for (let row = range.s.r + 1; row <= range.e.r; row += 1) {
    for (let column = range.s.c; column <= range.e.c; column += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r: row, c: column })];
      if (!cell || cell.t !== "n") continue;
      cell.z = /TON|小时数/.test(headers[column]) ? "#,##0.0" : "#,##0.###";
    }
  }
}

async function loadXlsxLibrary() {
  if (window.XLSX?.utils && typeof window.XLSX.write === "function") return window.XLSX;
  if (!state.xlsxModulePromise) {
    state.xlsxModulePromise = import("./vendor/xlsx.mjs?v=0.20.3").then((module) => {
      if (!module?.utils || typeof module.write !== "function") {
        throw new Error("SheetJS library is unavailable");
      }
      return module;
    });
  }
  return state.xlsxModulePromise;
}

function downloadWorkbook(XLSX, workbook, fileName) {
  const bytes = XLSX.write(workbook, { bookType: "xlsx", type: "array", compression: true });
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.hidden = true;
  document.body.append(link);
  link.click();
  link.remove();
  window.setTimeout(() => URL.revokeObjectURL(url), 30000);
}

async function exportTransferWorkbook(scope) {
  const XLSX = await loadXlsxLibrary();
  const selections = exportTransferSelection(scope);
  if (!selections.length) {
    throw new Error("当前视图没有可导出的中转记录");
  }
  const queryVvd = normalizedKey(els.vvdInput.value);
  const queryLane = normalizedKey(els.laneInput.value);
  const routeRows = [];
  const connectionRows = [];
  const blRows = [];
  for (const { group, segment } of selections) {
    routeRows.push({
      "查询 LANE": queryLane,
      "查询 VVD": queryVvd,
      "中转方向": transferDirectionLabel(segment.kind),
      "T/S PORT": group.port,
      "当前航段 POL": segment.start,
      "当前航段 POD": segment.end,
      "Booking 20'": segment.totals.b20,
      "Booking 40'": segment.totals.b40,
      "Booking TEU": segment.totals.bTeu,
      "OP TEU": segment.totals.opTeu,
      "VL TEU": segment.totals.vlTeu,
      "Weight TON": excelNumber(segment.totals.weight),
      "风险筛选": state.transferRisk === "all" || scope === "all" ? "全部" : transferRiskLabel(state.transferRisk),
      "T/S 类型筛选": state.transferType === "all" || scope === "all" ? "全部" : transferTypeLabel(state.transferType),
      "前程连接数": segment.connections.length,
      "BL 数": segment.count,
    });
    for (const connection of segment.connections) {
      connectionRows.push({
        "查询 LANE": queryLane,
        "查询 VVD": queryVvd,
        "中转方向": transferDirectionLabel(connection.kind),
        "T/S PORT": group.port,
        "当前航段": `${segment.start} → ${segment.end}`,
        "T/S 类型": transferTypeLabel(connection.transferType),
        "接驳状态": transferRiskLabel(connection.risk),
        "前段 POL": normalizedKey(connection.arrivalLeg.pol),
        "前段 LANE": normalizedKey(connection.arrivalLeg.lane),
        "前段 VVD": normalizedKey(connection.arrivalLeg.vvd),
        "前段 ETB": connection.timing.arrivalCall?.etb || "",
        "后段 LANE": normalizedKey(connection.departureLeg.lane),
        "后段 VVD": normalizedKey(connection.departureLeg.vvd),
        "后段 ETB": connection.timing.departureCall?.etb || "",
        "接驳小时数": Number.isFinite(connection.timing.gapHours) ? excelNumber(connection.timing.gapHours) : "",
        "Booking 20'": connection.totals.b20,
        "Booking 40'": connection.totals.b40,
        "Booking TEU": connection.totals.bTeu,
        "OP TEU": connection.totals.opTeu,
        "VL TEU": connection.totals.vlTeu,
        "Weight TON": excelNumber(connection.totals.weight),
        "BL 数": connection.recordIds.size,
      });
      for (const record of connection.records) {
        blRows.push({
          "BL No.": record.blNo || "",
          "CUL CODE": record.culCode || "",
          "查询 LANE": queryLane,
          "查询 VVD": queryVvd,
          "中转方向": transferDirectionLabel(connection.kind),
          "T/S PORT": group.port,
          "当前航段": `${segment.start} → ${segment.end}`,
          "T/S 类型": transferTypeLabel(connection.transferType),
          "接驳状态": transferRiskLabel(connection.risk),
          "前段 POL": normalizedKey(connection.arrivalLeg.pol),
          "前段 LANE": normalizedKey(connection.arrivalLeg.lane),
          "前段 VVD": normalizedKey(connection.arrivalLeg.vvd),
          "前段 ETB": connection.timing.arrivalCall?.etb || "",
          "后段 LANE": normalizedKey(connection.departureLeg.lane),
          "后段 VVD": normalizedKey(connection.departureLeg.vvd),
          "后段 ETB": connection.timing.departureCall?.etb || "",
          "接驳小时数": Number.isFinite(connection.timing.gapHours) ? excelNumber(connection.timing.gapHours) : "",
          "Booking 20'": record.b.t20,
          "Booking 40'": record.b.t40,
          "Booking TEU": record.b.teu,
          "OP TEU": record.op.teu,
          "VL TEU": record.vl.teu,
          "Weight TON": excelNumber(record.w),
          "COC/SOC": record.coc,
          "CUL/SUL": record.sul,
          "POR": record.original?.por || "",
          "POL": record.original?.pol || "",
          "POD": record.original?.pod || "",
          "DEL": record.original?.del || "",
          "完整路径": recordPath(record),
        });
      }
    }
  }
  const workbook = XLSX.utils.book_new();
  const routeSheet = XLSX.utils.json_to_sheet(routeRows);
  const connectionSheet = XLSX.utils.json_to_sheet(connectionRows);
  const blSheet = XLSX.utils.json_to_sheet(blRows);
  setWorksheetWidths(XLSX, routeSheet, [14, 16, 16, 12, 14, 14, 13, 13, 14, 11, 11, 14, 12, 18, 12, 10]);
  setWorksheetWidths(XLSX, connectionSheet, [14, 16, 16, 12, 22, 18, 13, 12, 12, 16, 18, 12, 16, 18, 13, 13, 13, 14, 11, 11, 14, 10]);
  setWorksheetWidths(XLSX, blSheet, [22, 16, 14, 16, 16, 12, 22, 18, 13, 12, 12, 16, 18, 12, 16, 18, 13, 13, 13, 14, 11, 11, 14, 12, 12, 12, 12, 12, 12, 48]);
  XLSX.utils.book_append_sheet(workbook, routeSheet, "Route Summary");
  XLSX.utils.book_append_sheet(workbook, connectionSheet, "Connection Detail");
  XLSX.utils.book_append_sheet(workbook, blSheet, "BL Detail");
  const scopeName = scope === "all" ? "ALL_TS" : scope === "segment" ? "SEGMENT" : "CURRENT_VIEW";
  const fileName = `${queryVvd || "VVD"}_${scopeName}.xlsx`;
  downloadWorkbook(XLSX, workbook, fileName);
  return fileName;
}

function initializeTransferMultiSelects() {
  els.transferMultiFilters.querySelectorAll("[data-multi-field]").forEach((container) => {
    const field = container.dataset.multiField;
    const label = TRANSFER_SEARCH_FIELDS[field] || container.dataset.label || field;
    container.innerHTML = `
      <span class="multi-select-label">${escapeHtml(label)}</span>
      <div class="multi-select-control" data-multi-control>
        <div class="multi-select-chips"></div>
        <input type="text" autocomplete="off" aria-label="${escapeHtml(label)}" placeholder="输入或选择，可多选" />
        <span class="multi-select-chevron" aria-hidden="true">⌄</span>
      </div>
      <div class="multi-select-menu" role="listbox" hidden></div>`;
  });
}

function renderTransferMultiSelect(field, open = false) {
  const container = els.transferMultiFilters.querySelector(`[data-multi-field="${field}"]`);
  if (!container) return;
  const selected = state.transferSearchFilters[field] || new Set();
  const chips = container.querySelector(".multi-select-chips");
  const input = container.querySelector("input");
  const menu = container.querySelector(".multi-select-menu");
  chips.innerHTML = [...selected]
    .sort(compareAlpha)
    .map((value) => `<button type="button" class="multi-select-chip" data-multi-remove="${escapeHtml(value)}"><span>${escapeHtml(value)}</span><em aria-hidden="true">×</em></button>`)
    .join("");
  const search = normalizedKey(input.value);
  const options = transferSearchFilterOptions(field)
    .filter((value) => !selected.has(value) && (!search || value.includes(search)))
    .slice(0, 160);
  menu.innerHTML = options.length
    ? options.map((value) => `<button type="button" role="option" data-multi-add="${escapeHtml(value)}">${escapeHtml(value)}</button>`).join("")
    : `<span class="multi-select-empty">${normalizedKey(els.transferPortInput.value) ? "没有其他候选" : "请先选择中转港"}</span>`;
  menu.hidden = !open;
  input.setAttribute("aria-expanded", open ? "true" : "false");
  container.classList.toggle("open", open);
}

function renderAllTransferMultiSelects() {
  Object.keys(TRANSFER_SEARCH_FIELDS).forEach((field) => renderTransferMultiSelect(field));
}

function commitTransferMultiChoice(target) {
  const container = target.closest("[data-multi-field]");
  if (!container) return false;
  const addButton = target.closest("[data-multi-add]");
  const removeButton = target.closest("[data-multi-remove]");
  if (!addButton && !removeButton) return false;

  const field = container.dataset.multiField;
  if (addButton) state.transferSearchFilters[field].add(normalizedKey(addButton.dataset.multiAdd));
  if (removeButton) state.transferSearchFilters[field].delete(normalizedKey(removeButton.dataset.multiRemove));
  const input = container.querySelector("input");
  if (input) input.value = "";
  markTransferSearchDirty();
  saveQuery();
  render();
  return true;
}

function transferMetricSummary(totals) {
  return `
    <span class="ts-metric booking"><small>Booking</small><strong>${display(totals.b20)} / ${display(totals.b40)}</strong><em>${display(totals.bTeu)} TEU</em></span>
    <span class="ts-metric op"><small>OP</small><strong>${display(totals.op20)} / ${display(totals.op40)}</strong><em>${display(totals.opTeu)} TEU</em></span>
    <span class="ts-metric vl"><small>VL</small><strong>${display(totals.vl20)} / ${display(totals.vl40)}</strong><em>${display(totals.vlTeu)} TEU</em></span>
    <span class="ts-metric weight"><small>Weight</small><strong>${display(totals.weight)}</strong><em>TON</em></span>`;
}

function transferSearchRiskStatus(events) {
  const priority = ["missed", "tight", "unknown", "normal"];
  const risk = priority.find((value) => events.some((event) => event.risk === value)) || "unknown";
  const affected = risk === "normal" ? events : events.filter((event) => event.risk === risk);
  return { risk, events: affected.length, totals: totalsForTransferEvents(affected) };
}

function compactTransferValues(values, limit = 6) {
  const sorted = [...new Set(values.map(normalizedKey).filter((value) => value && value !== "—"))].sort(compareAlpha);
  if (sorted.length <= limit) return sorted.join(" · ") || "—";
  return `${sorted.slice(0, limit).join(" · ")}　+${display(sorted.length - limit, true)}`;
}

function groupTransferResolvedEvents(events) {
  const outboundGroups = new Map();
  for (const event of events) {
    const outboundKey = event.vvdOut || "—";
    if (!outboundGroups.has(outboundKey)) {
      outboundGroups.set(outboundKey, {
        key: outboundKey,
        vvdOut: event.vvdOut,
        laneOuts: new Set(),
        ports: new Set(),
        destinations: new Set(),
        arrivalDates: new Set(),
        bookingPols: new Set(),
        events: [],
        groups: new Map(),
      });
    }
    const outboundGroup = outboundGroups.get(outboundKey);
    outboundGroup.laneOuts.add(event.laneOut);
    outboundGroup.ports.add(event.port);
    outboundGroup.destinations.add(event.to);
    outboundGroup.arrivalDates.add(event.arrivalDate);
    outboundGroup.bookingPols.add(event.bookingPol);
    outboundGroup.events.push(event);
    const inboundKey = joinKey(event.arrivalDate, event.vvdIn, event.from, event.to);
    if (!outboundGroup.groups.has(inboundKey)) {
      outboundGroup.groups.set(inboundKey, {
        key: inboundKey,
        arrivalDate: event.arrivalDate,
        vvdIn: event.vvdIn,
        laneIn: event.laneIn,
        vvdOut: event.vvdOut,
        laneOut: event.laneOut,
        from: event.from,
        port: event.port,
        to: event.to,
        events: [],
      });
    }
    outboundGroup.groups.get(inboundKey).events.push(event);
  }
  return [...outboundGroups.values()]
    .map((outboundGroup) => ({
      ...outboundGroup,
      laneOut: [...outboundGroup.laneOuts].sort(compareAlpha).join(" · "),
      port: [...outboundGroup.ports].sort(compareAlpha).join(" · "),
      to: [...outboundGroup.destinations].sort(compareAlpha).join(" · "),
      arrivalDateCount: outboundGroup.arrivalDates.size,
      bookingPol: compactTransferValues([...outboundGroup.bookingPols]),
      totals: totalsForTransferEvents(outboundGroup.events),
      riskStatus: transferSearchRiskStatus(outboundGroup.events),
      groups: [...outboundGroup.groups.values()]
        .map((group) => ({
          ...group,
          bookingPol: compactTransferValues(group.events.map((event) => event.bookingPol), 4),
          totals: totalsForTransferEvents(group.events),
          riskStatus: transferSearchRiskStatus(group.events),
        }))
        .sort((left, right) => left.arrivalDate.localeCompare(right.arrivalDate) || compareAlpha(left.vvdIn, right.vvdIn) || compareAlpha(left.from, right.from)),
    }))
    .sort((left, right) => compareAlpha(left.vvdOut, right.vvdOut));
}

function transferBlRows(events) {
  return events
    .sort((left, right) => compareAlpha(left.record.blNo, right.record.blNo))
    .map((event) => `
      <tr>
        <td><strong>${escapeHtml(event.record.blNo || "—")}</strong><small>${escapeHtml(event.record.culCode || "—")}</small></td>
        <td>${escapeHtml(event.bookingPol)} → ${escapeHtml(event.bookingPod)}</td>
        <td>${escapeHtml(event.laneIn)} / ${escapeHtml(event.vvdIn)}</td>
        <td>${escapeHtml(event.laneOut)} / ${escapeHtml(event.vvdOut)}</td>
        <td>${escapeHtml(event.record.coc || "—")} / ${escapeHtml(event.record.sul || "—")}</td>
        <td>${display(event.record.b.t20)} / ${display(event.record.b.t40)} / ${display(event.record.b.teu)}</td>
        <td>${display(event.record.op.t20)} / ${display(event.record.op.t40)} / ${display(event.record.op.teu)}</td>
        <td>${display(event.record.vl.t20)} / ${display(event.record.vl.t40)} / ${display(event.record.vl.teu)}</td>
      </tr>`)
    .join("");
}

function loadTransferBlDetails(details) {
  if (!details?.open || details.dataset.blLoaded === "true") return;
  const events = state.transferSearchGroupEvents.get(details.dataset.tsGroupId) || [];
  const placeholder = details.querySelector("[data-ts-bl-placeholder]");
  if (!placeholder) return;
  placeholder.innerHTML = `
    <table class="ts-bl-table">
      <thead><tr><th>BL No. / CUL CODE</th><th>Booking POL → POD</th><th>LANE / VVD IN</th><th>LANE / VVD OUT</th><th>箱源 / 货源</th><th>Booking 20/40/TEU</th><th>OP 20/40/TEU</th><th>VL 20/40/TEU</th></tr></thead>
      <tbody>${transferBlRows(events)}</tbody>
    </table>`;
  details.dataset.blLoaded = "true";
}

function transferResolvedGroupHtml(group, groupId) {
  const timing = group.events[0]?.timing || {};
  const risk = group.riskStatus.risk;
  const riskMeta = TRANSFER_RISK_META[risk] || TRANSFER_RISK_META.unknown;
  const multiCallKey = group.events.find((event) => event.arrivalCalls.length > 1)?.callKey || "";
  return `
    <details class="ts-route-group ts-connection-group risk-${escapeHtml(risk)}" data-ts-group-id="${escapeHtml(groupId)}">
      <summary>
        <span class="ts-arrival-identity"><small>VVD IN · ${escapeHtml(group.laneIn)}</small><strong>${escapeHtml(group.vvdIn)}</strong><em>${escapeHtml(group.from)} → ${escapeHtml(group.port)}　·　POL ${escapeHtml(group.bookingPol)}</em></span>
        <span class="ts-connection-times">
          <span><small>前程到港 ETB</small><strong>${escapeHtml(formatDate(timing.arrivalCall?.etb, true))}</strong></span>
          <i>→</i>
          <span><small>下一程离港 ETD</small><strong>${escapeHtml(formatDate(timing.departureCall?.etd, true))}</strong></span>
          <span class="ts-connection-gap ${escapeHtml(risk)}"><small>接驳</small><strong>${escapeHtml(formatEtbGap(timing.gapHours))}</strong></span>
        </span>
        <span class="ts-connection-cargo"><strong>${display(group.totals.bTeu)}</strong><small>TEU</small><em>${display(group.events.length, true)} BL</em></span>
        <span class="ts-connection-status ${escapeHtml(risk)}"><strong>${escapeHtml(riskMeta.short)}</strong>${multiCallKey ? `<button type="button" data-call-reset="${escapeHtml(multiCallKey)}">修改靠泊</button>` : ""}</span>
        <span class="ts-connection-expand"><span>查看 BL</span><span>收起 BL</span></span>
      </summary>
      <div class="ts-bl-table-wrap" data-ts-bl-placeholder><span class="ts-bl-loading">展开后载入 ${display(group.events.length, true)} 条 BL 明细</span></div>
    </details>`;
}

function transferOutboundRiskHtml(status) {
  const meta = TRANSFER_RISK_META[status.risk] || TRANSFER_RISK_META.unknown;
  if (status.risk === "normal") return `<span class="ts-outbound-risk normal"><strong>全部可接</strong></span>`;
  return `<span class="ts-outbound-risk ${escapeHtml(status.risk)}"><strong>${escapeHtml(meta.label)}</strong><em>${display(status.totals.bTeu)} TEU</em></span>`;
}

function renderTransferResolvedResults(events) {
  const outboundGroups = groupTransferResolvedEvents(events);
  state.transferSearchGroupEvents.clear();
  els.transferSearchResults.innerHTML = outboundGroups.length
    ? outboundGroups.map((outboundGroup) => `
        <details class="ts-outbound-group risk-${escapeHtml(outboundGroup.riskStatus.risk)}">
          <summary class="ts-outbound-summary">
            <span class="ts-outbound-identity"><small>VVD OUT · ${escapeHtml(outboundGroup.laneOut || "—")}</small><strong>${escapeHtml(outboundGroup.vvdOut || "—")}</strong><em>${escapeHtml(outboundGroup.port)} → ${escapeHtml(outboundGroup.to)}</em></span>
            <span class="ts-outbound-cargo"><small>中转货量</small><strong>${display(outboundGroup.totals.bTeu)}</strong><em>TEU　·　${display(outboundGroup.events.length, true)} BL</em></span>
            <span class="ts-outbound-pols"><small>来自 Booking POL</small><strong>${escapeHtml(outboundGroup.bookingPol)}</strong><em>${display(outboundGroup.arrivalDateCount, true)} 个 ETB 日期</em></span>
            ${transferOutboundRiskHtml(outboundGroup.riskStatus)}
            <span class="ts-outbound-toggle"><span>展开接转详情</span><span>收起接转详情</span></span>
          </summary>
          <div class="ts-outbound-details">${outboundGroup.groups.map((group) => {
            const groupId = `${outboundGroup.key}|${group.key}`;
            state.transferSearchGroupEvents.set(groupId, group.events);
            return transferResolvedGroupHtml(group, groupId);
          }).join("")}</div>
        </details>`).join("")
    : `<div class="module-empty"><span class="module-empty-mark">T/S</span><div><strong>当前日期与筛选条件没有已确认到港记录</strong><p>可调整日期、选填条件，或先处理上方“待选择靠泊”。</p></div></div>`;
}

function renderTransferCallChoices(pending) {
  els.transferCallChoiceSection.hidden = !pending.length;
  els.transferCallChoices.innerHTML = pending.map((group) => {
    const totals = totalsForTransferEvents(group.events);
    return `
      <article class="transfer-call-choice-card">
        <div class="call-choice-identity"><span>多靠泊航次</span><strong>${escapeHtml(group.event.vvdIn)}</strong><small>${escapeHtml(group.event.laneIn)} · ${escapeHtml(group.event.from)} → ${escapeHtml(group.event.port)} · ${display(group.events.length, true)} BL</small></div>
        <div class="call-choice-options">
          ${group.calls.map((call, index) => `
            <button type="button" data-call-key="${escapeHtml(group.key)}" data-call-signature="${escapeHtml(scheduleCallSignature(call))}">
              <span>靠泊 ${index + 1}</span><strong>ETB ${escapeHtml(formatDate(call.etb))}</strong><small>ETA ${escapeHtml(formatDate(call.eta))}　ETD ${escapeHtml(formatDate(call.etd))}</small>
            </button>`).join("")}
        </div>
        <div class="call-choice-metrics">${transferMetricSummary(totals)}</div>
      </article>`;
  }).join("");
}

function renderTransferMissingResults(events) {
  els.transferMissingSection.hidden = !events.length;
  state.transferMissingGroupEvents.clear();
  const groups = new Map();
  for (const event of events) {
    const key = joinKey(event.from, event.port, event.to, event.laneIn, event.vvdIn, event.laneOut, event.vvdOut);
    if (!groups.has(key)) groups.set(key, { event, events: [] });
    groups.get(key).events.push(event);
  }
  els.transferMissingResults.innerHTML = [...groups.values()]
    .sort((left, right) => right.events.length - left.events.length)
    .map(({ event, events: groupEvents }) => {
      const totals = totalsForTransferEvents(groupEvents);
      const groupId = `missing|${event.port}|${event.from}|${event.to}|${event.laneIn}|${event.vvdIn}|${event.laneOut}|${event.vvdOut}`;
      state.transferMissingGroupEvents.set(groupId, groupEvents);
      return `<details class="missing-route-group" data-missing-group-id="${escapeHtml(groupId)}"><summary><span><strong>${escapeHtml(event.from)} → ${escapeHtml(event.port)} → ${escapeHtml(event.to)}</strong><small>${escapeHtml(event.laneIn)} / ${escapeHtml(event.vvdIn)}　→　${escapeHtml(event.laneOut)} / ${escapeHtml(event.vvdOut)}</small></span><em>${display(groupEvents.length, true)} BL</em><span class="missing-route-metrics">${display(totals.bTeu)} TEU</span></summary><div class="missing-bl-list" data-missing-bl-placeholder><span class="missing-bl-loading">展开后载入 ${display(groupEvents.length, true)} 条待匹配 BL</span></div></details>`;
    }).join("");
}

function loadTransferMissingDetails(details) {
  if (!details?.open || details.dataset.blLoaded === "true") return;
  const events = state.transferMissingGroupEvents.get(details.dataset.missingGroupId) || [];
  const placeholder = details.querySelector("[data-missing-bl-placeholder]");
  if (!placeholder) return;
  placeholder.innerHTML = events.map((event) => `<span>${escapeHtml(event.record.blNo || "—")}</span>`).join("");
  details.dataset.blLoaded = "true";
}

function setTransferSearchTotals(totals) {
  els.transferTotalB20.textContent = display(totals.b20);
  els.transferTotalB40.textContent = display(totals.b40);
  els.transferTotalBTeu.textContent = display(totals.bTeu);
  els.transferTotalOp20.textContent = display(totals.op20);
  els.transferTotalOp40.textContent = display(totals.op40);
  els.transferTotalOpTeu.textContent = display(totals.opTeu);
  els.transferTotalVl20.textContent = display(totals.vl20);
  els.transferTotalVl40.textContent = display(totals.vl40);
  els.transferTotalVlTeu.textContent = display(totals.vlTeu);
  els.transferTotalWeight.textContent = display(totals.weight);
}

function renderTransferSearch() {
  if (!state.transferSearchSubmitted) {
    els.transferSearchHint.classList.remove("warning");
    els.transferSearchHint.textContent = "选择中转港，并填写 ETB 日期或至少一项下方查询条件";
    showEmpty("设置条件后点击查询", "中转港为必填；填写 ETB 开始或结束日期后，下方条件可以全部留空。日期全部留空时，下方至少选择一项。", true);
    return;
  }
  const view = buildTransferSearchView();
  const portExists = state.transferEventsByPort.has(view.port);
  els.transferSearchHint.classList.remove("warning");
  if (!view.port) {
    els.transferSearchHint.textContent = "请选择中转港";
    els.transferSearchHint.classList.add("warning");
    showEmpty("中转港不能为空", "请输入或从候选列表选择一个中转港。", true);
    return;
  }
  if (!hasTransferSearchScope(view.start, view.end)) {
    els.transferSearchHint.textContent = "请填写 ETB 日期或至少一项下方查询条件";
    els.transferSearchHint.classList.add("warning");
    showEmpty("缺少查询范围", "请填写 ETB 开始或结束日期；如果日期留空，请在 Pre/Post Port、Booking POL/POD、LANE IN/OUT 或 VVD IN/OUT 中至少选择一个值。", true);
    return;
  }
  if (!validTransferDate(view.start) || !validTransferDate(view.end)) {
    els.transferSearchHint.textContent = "日期格式应为 YYYY-MM-DD，例如 2026-06-01";
    els.transferSearchHint.classList.add("warning");
    showEmpty("日期格式不正确", "可留空；如填写，请输入 2026-06-01，或直接输入 20260601 让系统自动转换。", true);
    return;
  }
  if (!portExists) {
    els.transferSearchHint.textContent = `没有找到中转港 ${view.port}`;
    els.transferSearchHint.classList.add("warning");
    showEmpty("没有找到这个中转港", "请从候选列表选择一个由相邻两程形成的中转港。", true);
    return;
  }
  if (view.start && view.end && view.start > view.end) {
    els.transferSearchHint.textContent = "结束日期不能早于开始日期";
    els.transferSearchHint.classList.add("warning");
    showEmpty("日期区段无效", "请调整开始日期或结束日期。", true);
    return;
  }

  const totals = totalsForTransferEvents(view.resolved);
  const pendingTotals = totalsForTransferEvents(view.pending.flatMap((group) => group.events));
  const missingTotals = totalsForTransferEvents(view.missing);
  const uniqueBlCount = new Set(view.resolved.map((event) => event.record._id)).size;
  const dateScope = view.start && view.end
    ? `ETB ${view.start} 至 ${view.end}（包含首尾两天）`
    : view.start
      ? `ETB ${view.start} 起`
      : view.end
        ? `ETB 截至 ${view.end}`
        : "ETB 不限日期";
  els.transferSearchSummaryTitle.textContent = `${view.port} 中转货量汇总`;
  els.transferSearchSummaryHint.textContent = `${dateScope}　${uniqueBlCount} 个唯一 BL / ${view.resolved.length} 个中转事件`;
  els.transferSearchMatchedCount.textContent = display(totals.bTeu);
  els.transferSearchPendingCount.textContent = display(pendingTotals.bTeu);
  els.transferSearchMissingCount.textContent = display(missingTotals.bTeu);
  els.transferSearchSwitchTeu.textContent = display(totals.bTeu);
  els.transferSearchSwitchWeight.textContent = display(totals.weight);
  setTransferSearchTotals(totals);
  renderTransferCallChoices(view.pending);
  renderTransferResolvedResults(view.resolved);
  renderTransferMissingResults(view.missing);
  els.emptyState.hidden = true;
  els.resultContent.hidden = true;
  els.transferSearchContent.hidden = false;
}

function renderOverview(matches, directMatches, transferMatches, directGroups, transferGroups) {
  const lane = normalizedKey(els.laneInput.value);
  const vvd = normalizedKey(els.vvdInput.value);
  const coc = filterValue(els.cocFilter);
  const sul = filterValue(els.sulFilter);
  const filters = [coc !== "ALL" ? coc : "COC + SOC", sul !== "ALL" ? sul : "CUL + SUL"];
  const matchTotals = totalsForMatches(matches).totals;
  const directTotals = totalsForMatches(directMatches).totals;
  const transferTotals = totalsForMatches(transferMatches).totals;

  els.queryLane.textContent = lane || "LANE";
  els.overviewTitle.textContent = vvd;
  els.querySummary.textContent = `匹配 1st–5th leg　${filters.join("　")}`;
  els.matchTeu.textContent = display(matchTotals.bTeu);
  els.matchWeight.textContent = display(matchTotals.weight);
  els.directTeu.textContent = display(directTotals.bTeu);
  els.directWeight.textContent = display(directTotals.weight);
  els.tsTeu.textContent = display(transferTotals.bTeu);
  els.tsWeight.textContent = display(transferTotals.weight);
  els.directSwitchTeu.textContent = display(directTotals.bTeu);
  els.directSwitchWeight.textContent = display(directTotals.weight);
  els.transferSwitchTeu.textContent = display(transferTotals.bTeu);
  els.transferSwitchWeight.textContent = display(transferTotals.weight);
  els.directBadge.textContent = `${display(directGroups.length, true)} POL`;
  els.transferBadge.textContent = `${display(transferGroups.length, true)} T/S 港`;
}

function renderModuleSummary(directMatches, transferMatches) {
  const direct = state.activeModule === "direct";
  const { totals, count } = totalsForMatches(direct ? directMatches : transferMatches);
  const transferScope = state.transferDirection === "inbound"
    ? "转入当前 VVD"
    : state.transferDirection === "outbound"
      ? "转出当前 VVD"
      : "全部中转衔接";
  els.moduleSummaryKicker.textContent = direct ? "DIRECT MONITOR" : "TRANSIT MONITOR";
  els.moduleSummaryTitle.textContent = direct ? "直航货量汇总" : "中转货量汇总";
  els.moduleSummaryHint.textContent = direct
    ? `仅统计当前直航模块的 ${display(count, true)} 条唯一 Booking`
    : `当前查看：${transferScope}　${display(count, true)} 条唯一 BL，不重复累计中转港`;
  els.totalWeight.textContent = display(totals.weight);
  els.totalB20.textContent = display(totals.b20);
  els.totalB40.textContent = display(totals.b40);
  els.totalBTeu.textContent = display(totals.bTeu);
  els.totalOp20.textContent = display(totals.op20);
  els.totalOp40.textContent = display(totals.op40);
  els.totalOpTeu.textContent = display(totals.opTeu);
  els.totalVl20.textContent = display(totals.vl20);
  els.totalVl40.textContent = display(totals.vl40);
  els.totalVlTeu.textContent = display(totals.vlTeu);
}

function renderTransferDirection(transferMatches) {
  const labels = {
    all: "同时查看转入与转出；中间程 Booking 可在两个方向分别出现。",
    inbound: "查看由其他航线经中转港转入当前搜索 VVD 的 Booking。",
    outbound: "查看由当前搜索 VVD 经中转港转往其他航线的 Booking。",
  };
  els.transferDirectionHint.textContent = labels[state.transferDirection];
  els.transferDirection.querySelectorAll("[data-transfer-direction]").forEach((button) => {
    const direction = button.dataset.transferDirection;
    const active = direction === state.transferDirection;
    button.classList.toggle("active", active);
    button.setAttribute("aria-pressed", active ? "true" : "false");
    const count = totalsForMatches(filterTransferMatches(transferMatches, direction)).count;
    const countLabel = button.querySelector("[data-transfer-count]");
    if (countLabel) countLabel.textContent = `${display(count, true)} 条`;
  });
}

function updateModuleChrome() {
  els.moduleSwitch.querySelectorAll("[data-module]").forEach((button) => {
    const active = button.dataset.module === state.activeModule;
    button.classList.toggle("active", active);
    button.setAttribute("aria-pressed", active ? "true" : "false");
  });
  const transferSearch = state.activeModule === "transfer-search";
  const direct = state.activeModule === "direct";
  els.voyageSearchWorkbench.hidden = transferSearch;
  els.transferSearchWorkbench.hidden = !transferSearch;
  els.resultContent.dataset.module = state.activeModule;
  els.directPanel.hidden = transferSearch || !direct;
  els.transferPanel.hidden = transferSearch || direct;
  els.directPanel.classList.toggle("active", !transferSearch && direct);
  els.transferPanel.classList.toggle("active", !transferSearch && !direct);
  if (direct || transferSearch) document.body.classList.remove("transfer-modal-open");
}

function showEmpty(title, text, transferSearch = false) {
  els.emptyTitle.textContent = title;
  els.emptyText.textContent = text;
  els.emptyState.dataset.context = transferSearch ? "transfer-search" : "voyage";
  els.emptyState.hidden = false;
  els.resultContent.hidden = true;
  els.transferSearchContent.hidden = true;
}

function showResults() {
  els.emptyState.hidden = true;
  els.resultContent.hidden = false;
  els.transferSearchContent.hidden = true;
}

function renderQueryHint() {
  const lane = normalizedKey(els.laneInput.value);
  const vvd = normalizedKey(els.vvdInput.value);
  els.queryHint.classList.remove("warning");
  if (!vvd) {
    els.queryHint.textContent = lane
      ? `已按 ${lane} 收窄 VVD 候选，请选择一个具体 VVD`
      : "VVD 是主查询条件，输入后会自动匹配 Lane";
    return;
  }
  const lanes = state.data?.vvdLanes?.[vvd];
  if (!lanes) {
    els.queryHint.textContent = `没有找到 ${vvd}`;
    els.queryHint.classList.add("warning");
    return;
  }
  els.queryHint.textContent = `${vvd} 已匹配 ${lane || lanes[0]?.lane || "对应"} Lane`;
}

function render() {
  if (!state.data) return;
  document.body.classList.remove("transfer-modal-open");
  updateModuleChrome();
  if (state.activeModule === "transfer-search") {
    renderAllTransferMultiSelects();
    renderTransferSearch();
    return;
  }
  renderQueryHint();
  const lane = normalizedKey(els.laneInput.value);
  const vvd = normalizedKey(els.vvdInput.value);
  if (!vvd) {
    showEmpty(
      lane ? "请选择一个具体 VVD" : "输入 VVD 开始监控",
      lane
        ? `Lane ${lane} 只用于收窄候选；为避免混合多个航次，请继续选择 VVD。`
        : "系统会从 1st–5th leg 中定位航次，并分别显示直航与中转衔接。",
    );
    return;
  }

  const matches = collectMatches();
  if (!matches.length) {
    showEmpty(
      "没有匹配结果",
      "请检查 VVD 是否完整，或将 COC/SOC、CUL/SUL 筛选切回“全部”后重试。",
    );
    return;
  }

  const directMatches = matches.filter(({ record }) => record.legs.length === 1);
  const transferMatches = matches.filter(({ record }) => record.legs.length > 1);
  const directGroups = buildDirectGroups(directMatches);
  const allTransferGroups = buildTransferGroups(transferMatches, "all");
  const visibleTransferMatches = filterTransferMatches(transferMatches);
  const transferGroups = buildTransferGroups(transferMatches);
  const signature = [lane, vvd, filterValue(els.cocFilter), filterValue(els.sulFilter)].join("|");
  if (signature !== state.lastQuerySignature) {
    state.lastQuerySignature = signature;
    state.expandedPols.clear();
    state.expandedConnections.clear();
    state.expandedFeederGroups.clear();
    state.drawerSegmentKey = "";
    state.selectedTransferPort = "";
    state.selectedTransferSegment = "";
    if (!directMatches.length && transferMatches.length) state.activeModule = "transfer";
    if (!transferMatches.length && directMatches.length) state.activeModule = "direct";
  }

  renderOverview(matches, directMatches, transferMatches, directGroups, allTransferGroups);
  renderTransferDirection(transferMatches);
  renderModuleSummary(directMatches, visibleTransferMatches);
  if (state.activeModule === "direct") {
    renderDirect(directGroups);
  } else {
    renderTransferPortTabs(transferGroups);
    renderTransfer(transferGroups);
  }
  showResults();
}

function syncLaneFromVvd() {
  const vvd = normalizedKey(els.vvdInput.value);
  if (!vvd) return;
  const lanes = state.data.vvdLanes[vvd];
  if (!lanes) {
    els.laneInput.value = "";
    return;
  }
  els.laneInput.value = lanes[0]?.lane || "";
}

function refreshVvdList() {
  const lane = normalizedKey(els.laneInput.value);
  const source = lane && state.data.laneVvds[lane] ? state.data.laneVvds[lane] : state.data.vvds;
  els.vvdList.innerHTML = optionList(source);
}

function currentQuery() {
  return {
    lane: els.laneInput.value.trim(),
    vvd: els.vvdInput.value.trim(),
    coc: filterValue(els.cocFilter),
    sul: filterValue(els.sulFilter),
    module: state.activeModule,
    transferDirection: state.transferDirection,
    transferRisk: state.transferRisk,
    transferType: state.transferType,
    transferPort: els.transferPortInput.value.trim(),
    transferStartDate: els.transferStartDate.value,
    transferEndDate: els.transferEndDate.value,
    transferCoc: transferSearchFilterValue(els.transferCocFilter),
    transferSul: transferSearchFilterValue(els.transferSulFilter),
    transferFilters: Object.fromEntries(
      Object.entries(state.transferSearchFilters).map(([field, values]) => [field, [...values]]),
    ),
  };
}

function saveQuery() {
  try {
    const query = currentQuery();
    const hasTransferOptional = Object.values(query.transferFilters).some((values) => values.length);
    if (!query.lane && !query.vvd && !query.transferPort && !query.transferStartDate && !query.transferEndDate && query.coc === "ALL" && query.sul === "ALL" && query.transferCoc === "ALL" && query.transferSul === "ALL" && !hasTransferOptional) {
      localStorage.removeItem(QUERY_STORAGE_KEY);
      return;
    }
    localStorage.setItem(QUERY_STORAGE_KEY, JSON.stringify(query));
  } catch {
    // The monitor remains usable when browser storage is unavailable.
  }
}

function restoreQuery() {
  let query = null;
  try {
    query = JSON.parse(localStorage.getItem(QUERY_STORAGE_KEY) || "null");
  } catch {
    query = null;
  }
  if (!query) return;
  els.laneInput.value = query.lane || "";
  els.vvdInput.value = query.vvd || "";
  setFilterValue(els.cocFilter, query.coc || "ALL");
  setFilterValue(els.sulFilter, query.sul || "ALL");
  state.activeModule = ["transfer", "transfer-search"].includes(query.module) ? query.module : "direct";
  state.transferDirection = ["inbound", "outbound"].includes(query.transferDirection)
    ? query.transferDirection
    : "all";
  state.transferRisk = ["missed", "tight", "normal", "unknown"].includes(query.transferRisk)
    ? query.transferRisk
    : "all";
  state.transferType = ["domestic", "international", "unknown"].includes(query.transferType)
    ? query.transferType
    : "all";
  els.transferPortInput.value = query.transferPort || "";
  els.transferStartDate.value = query.transferStartDate || "";
  els.transferEndDate.value = query.transferEndDate || "";
  setFilterValue(els.transferCocFilter, query.transferCoc || "ALL");
  setFilterValue(els.transferSulFilter, query.transferSul || "ALL");
  for (const field of Object.keys(TRANSFER_SEARCH_FIELDS)) {
    state.transferSearchFilters[field] = new Set(
      Array.isArray(query.transferFilters?.[field]) ? query.transferFilters[field].map(normalizedKey).filter(Boolean) : [],
    );
  }
  if (els.vvdInput.value.trim()) syncLaneFromVvd();
}

function clearSavedQuery() {
  try {
    localStorage.removeItem(QUERY_STORAGE_KEY);
  } catch {
    // Ignore blocked-storage failures.
  }
}

function wireEvents() {
  els.refreshBookingButton?.addEventListener("click", () => {
    refreshBookingData({ manual: true });
  });

  els.laneInput.addEventListener("input", () => {
    if (els.vvdInput.value.trim()) syncLaneFromVvd();
    refreshVvdList();
    saveQuery();
    render();
  });

  els.vvdInput.addEventListener("input", () => {
    syncLaneFromVvd();
    refreshVvdList();
    saveQuery();
    render();
  });

  [els.laneInput, els.vvdInput].forEach((input) => {
    input.addEventListener("change", () => {
      if (input === els.vvdInput) syncLaneFromVvd();
      refreshVvdList();
      saveQuery();
      render();
    });
  });

  [els.cocFilter, els.sulFilter].forEach((group) => {
    group.addEventListener("click", (event) => {
      const button = event.target.closest(".filter-option");
      if (!button || !group.contains(button)) return;
      setFilterValue(group, button.dataset.value);
      saveQuery();
      render();
    });
  });

  els.clearButton.addEventListener("click", () => {
    els.laneInput.value = "";
    els.vvdInput.value = "";
    setFilterValue(els.cocFilter, "ALL");
    setFilterValue(els.sulFilter, "ALL");
    state.activeModule = "direct";
    state.transferDirection = "all";
    state.transferRisk = "all";
    state.transferType = "all";
    state.lastQuerySignature = "";
    state.expandedPols.clear();
    state.expandedConnections.clear();
    state.expandedFeederGroups.clear();
    state.drawerSegmentKey = "";
    state.selectedTransferPort = "";
    state.selectedTransferSegment = "";
    refreshVvdList();
    clearSavedQuery();
    render();
    els.vvdInput.focus();
  });

  els.moduleSwitch.addEventListener("click", (event) => {
    const button = event.target.closest("[data-module]");
    if (!button) return;
    state.activeModule = ["transfer", "transfer-search"].includes(button.dataset.module)
      ? button.dataset.module
      : "direct";
    saveQuery();
    render();
  });

  els.transferPortInput.addEventListener("input", () => {
    markTransferSearchDirty();
    renderAllTransferMultiSelects();
    saveQuery();
    render();
  });
  els.transferPortInput.addEventListener("change", () => {
    markTransferSearchDirty();
    renderAllTransferMultiSelects();
    saveQuery();
    render();
  });

  [els.transferStartDate, els.transferEndDate].forEach((input) => {
    input.addEventListener("input", () => {
      if (/^\d{8}$/.test(input.value.trim())) normalizeTransferDateInput(input);
      markTransferSearchDirty();
      saveQuery();
      render();
    });
    input.addEventListener("change", () => {
      normalizeTransferDateInput(input);
      markTransferSearchDirty();
      saveQuery();
      render();
    });
    input.addEventListener("blur", () => {
      normalizeTransferDateInput(input);
      saveQuery();
    });
    input.addEventListener("keydown", (event) => {
      if (event.key !== "Enter") return;
      event.preventDefault();
      els.transferSearchButton.click();
    });
  });

  els.transferSearchWorkbench.querySelectorAll("[data-date-picker-for]").forEach((button) => {
    const input = document.getElementById(button.dataset.datePickerFor);
    const picker = button.parentElement.querySelector(".date-picker-proxy");
    if (!input || !picker) return;
    button.addEventListener("click", () => {
      const current = normalizeTransferDateInput(input);
      picker.value = validTransferDate(current) ? current : "";
      if (typeof picker.showPicker === "function") {
        picker.showPicker();
      } else {
        picker.click();
      }
    });
    picker.addEventListener("change", () => {
      if (!picker.value) return;
      input.value = picker.value;
      markTransferSearchDirty();
      saveQuery();
      render();
      input.focus();
    });
  });

  [els.transferCocFilter, els.transferSulFilter].forEach((group) => {
    group.addEventListener("click", (event) => {
      const button = event.target.closest(".filter-option");
      if (!button || !group.contains(button)) return;
      setFilterValue(group, button.dataset.value);
      state.transferSearchSubmitted = true;
      renderAllTransferMultiSelects();
      saveQuery();
      render();
    });
  });

  els.transferSearchButton.addEventListener("click", () => {
    normalizeTransferDateInput(els.transferStartDate);
    normalizeTransferDateInput(els.transferEndDate);
    state.transferSearchSubmitted = true;
    saveQuery();
    render();
  });

  els.clearTransferSearchButton.addEventListener("click", () => {
    els.transferPortInput.value = "";
    els.transferStartDate.value = "";
    els.transferEndDate.value = "";
    els.transferSearchWorkbench.querySelectorAll(".date-picker-proxy").forEach((picker) => { picker.value = ""; });
    setFilterValue(els.transferCocFilter, "ALL");
    setFilterValue(els.transferSulFilter, "ALL");
    Object.keys(TRANSFER_SEARCH_FIELDS).forEach((field) => state.transferSearchFilters[field].clear());
    state.transferCallSelections.clear();
    markTransferSearchDirty();
    saveQuery();
    render();
    els.transferPortInput.focus();
  });

  els.clearTransferOptionalButton.addEventListener("click", () => {
    setFilterValue(els.transferCocFilter, "ALL");
    setFilterValue(els.transferSulFilter, "ALL");
    Object.keys(TRANSFER_SEARCH_FIELDS).forEach((field) => state.transferSearchFilters[field].clear());
    markTransferSearchDirty();
    renderAllTransferMultiSelects();
    saveQuery();
    render();
  });

  els.transferMultiFilters.addEventListener("focusin", (event) => {
    const container = event.target.closest("[data-multi-field]");
    if (!container) return;
    renderTransferMultiSelect(container.dataset.multiField, true);
  });

  els.transferMultiFilters.addEventListener("input", (event) => {
    const container = event.target.closest("[data-multi-field]");
    if (!container || event.target.tagName !== "INPUT") return;
    renderTransferMultiSelect(container.dataset.multiField, true);
  });

  els.transferMultiFilters.addEventListener("pointerdown", (event) => {
    const choice = event.target.closest("[data-multi-add], [data-multi-remove]");
    if (!choice) return;
    choice.dataset.pointerHandled = "true";
    event.preventDefault();
    event.stopPropagation();
    commitTransferMultiChoice(event.target);
  });

  els.transferMultiFilters.addEventListener("click", (event) => {
    // Physical mouse/touch selection is committed on pointerdown, before a focus
    // change can close the menu. Click remains as a keyboard/legacy fallback.
    const choice = event.target.closest("[data-multi-add], [data-multi-remove]");
    if (choice?.dataset.pointerHandled === "true") {
      delete choice.dataset.pointerHandled;
      return;
    }
    if (commitTransferMultiChoice(event.target)) {
      event.preventDefault();
      event.stopPropagation();
    }
  });

  els.transferSearchContent.addEventListener("click", (event) => {
    const choice = event.target.closest("[data-call-key][data-call-signature]");
    if (choice) {
      state.transferCallSelections.set(choice.dataset.callKey, choice.dataset.callSignature);
      render();
      return;
    }
    const reset = event.target.closest("[data-call-reset]");
    if (reset) {
      event.preventDefault();
      state.transferCallSelections.delete(reset.dataset.callReset);
      render();
    }
  });

  els.transferSearchResults.addEventListener("toggle", (event) => {
    const details = event.target.closest("details[data-ts-group-id]");
    if (details) loadTransferBlDetails(details);
  }, true);

  els.transferMissingResults.addEventListener("toggle", (event) => {
    const details = event.target.closest("details[data-missing-group-id]");
    if (details) loadTransferMissingDetails(details);
  }, true);

  els.transferDirection.addEventListener("click", (event) => {
    const button = event.target.closest("[data-transfer-direction]");
    if (!button) return;
    state.transferDirection = ["inbound", "outbound"].includes(button.dataset.transferDirection)
      ? button.dataset.transferDirection
      : "all";
    state.selectedTransferPort = "";
    state.selectedTransferSegment = "";
    state.expandedConnections.clear();
    state.expandedFeederGroups.clear();
    state.drawerSegmentKey = "";
    saveQuery();
    render();
  });

  els.directRows.addEventListener("click", (event) => {
    const button = event.target.closest("[data-pol-toggle]");
    if (!button) return;
    const key = button.dataset.polToggle;
    if (state.expandedPols.has(key)) state.expandedPols.delete(key);
    else state.expandedPols.add(key);
    renderDirect(buildDirectGroups(collectMatches()));
  });

  els.transferPortTabs.addEventListener("click", (event) => {
    const button = event.target.closest("[data-ts-port]");
    if (!button) return;
    state.selectedTransferPort = button.dataset.tsPort || "";
    state.selectedTransferSegment = "";
    state.expandedConnections.clear();
    state.expandedFeederGroups.clear();
    state.drawerSegmentKey = "";
    const groups = buildTransferGroups(collectMatches());
    renderTransferPortTabs(groups);
    renderTransfer(groups);
  });

  els.transferRows.addEventListener("click", async (event) => {
    const exportButton = event.target.closest("[data-export-scope]");
    if (exportButton) {
      const buttons = [...els.transferRows.querySelectorAll("[data-export-scope]")];
      const status = els.transferRows.querySelector(".transfer-export-status");
      const originalText = exportButton.textContent;
      buttons.forEach((button) => { button.disabled = true; });
      exportButton.textContent = "正在生成…";
      if (status) {
        status.textContent = "正在整理中转数据并生成 Excel…";
        status.dataset.state = "working";
      }
      try {
        const fileName = await exportTransferWorkbook(exportButton.dataset.exportScope || "current");
        if (status) {
          status.textContent = `${fileName} 已开始下载`;
          status.dataset.state = "success";
        }
      } catch (error) {
        console.error("Transfer Excel export failed", error);
        if (status) {
          status.textContent = `导出失败：${error?.message || "浏览器未能生成文件"}`;
          status.dataset.state = "error";
        }
      } finally {
        buttons.forEach((button) => { button.disabled = false; });
        exportButton.textContent = originalText;
      }
      return;
    }
    if (event.target.closest("[data-transfer-drawer-close]")) {
      state.selectedTransferSegment = "";
      state.expandedConnections.clear();
      state.expandedFeederGroups.clear();
      state.drawerSegmentKey = "";
      renderTransfer(buildTransferGroups(collectMatches()));
      return;
    }
    const riskButton = event.target.closest("[data-transfer-risk]");
    if (riskButton) {
      state.transferRisk = ["missed", "tight", "normal", "unknown"].includes(riskButton.dataset.transferRisk)
        ? riskButton.dataset.transferRisk
        : "all";
      state.selectedTransferSegment = "";
      state.expandedConnections.clear();
      state.expandedFeederGroups.clear();
      state.drawerSegmentKey = "";
      saveQuery();
      render();
      return;
    }
    const typeButton = event.target.closest("[data-transfer-type]");
    if (typeButton) {
      state.transferType = ["domestic", "international", "unknown"].includes(typeButton.dataset.transferType)
        ? typeButton.dataset.transferType
        : "all";
      state.selectedTransferSegment = "";
      state.expandedConnections.clear();
      state.expandedFeederGroups.clear();
      state.drawerSegmentKey = "";
      saveQuery();
      render();
      return;
    }
    const segmentButton = event.target.closest("[data-transfer-segment]");
    if (segmentButton) {
      state.selectedTransferSegment = segmentButton.dataset.transferSegment || "";
      state.expandedConnections.clear();
      state.expandedFeederGroups.clear();
      state.drawerSegmentKey = "";
      renderTransfer(buildTransferGroups(collectMatches()));
      return;
    }
    const feederGroupButton = event.target.closest("[data-feeder-group-toggle]");
    if (feederGroupButton) {
      const key = feederGroupButton.dataset.feederGroupToggle;
      const expanded = !state.expandedFeederGroups.has(key);
      if (expanded) state.expandedFeederGroups.add(key);
      else state.expandedFeederGroups.delete(key);
      feederGroupButton.setAttribute("aria-expanded", expanded ? "true" : "false");
      const body = document.getElementById(feederGroupButton.getAttribute("aria-controls"));
      if (body) body.hidden = !expanded;
      const group = feederGroupButton.closest(".feeder-pol-group");
      if (group) group.classList.toggle("expanded", expanded);
      const chevron = feederGroupButton.querySelector(".feeder-pol-chevron");
      if (chevron) chevron.textContent = expanded ? "−" : "+";
      return;
    }
    const connectionButton = event.target.closest("[data-connection-toggle]");
    if (connectionButton) {
      const key = connectionButton.dataset.connectionToggle;
      const expanded = !state.expandedConnections.has(key);
      if (expanded) state.expandedConnections.add(key);
      else state.expandedConnections.delete(key);
      connectionButton.setAttribute("aria-expanded", expanded ? "true" : "false");
      const details = document.getElementById(connectionButton.getAttribute("aria-controls"));
      if (details) details.hidden = !expanded;
      const label = connectionButton.querySelector(".feeder-expand");
      if (label) label.textContent = expanded ? "收起 BL" : "查看 BL";
    }
  });

  document.addEventListener("click", (event) => {
    els.transferMultiFilters.querySelectorAll("[data-multi-field]").forEach((container) => {
      if (!container.contains(event.target)) renderTransferMultiSelect(container.dataset.multiField, false);
    });
  });

  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    els.transferMultiFilters.querySelectorAll("[data-multi-field]").forEach((container) => {
      renderTransferMultiSelect(container.dataset.multiField, false);
    });
    if (!state.selectedTransferSegment) return;
    state.selectedTransferSegment = "";
    state.expandedConnections.clear();
    state.expandedFeederGroups.clear();
    state.drawerSegmentKey = "";
    renderTransfer(buildTransferGroups(collectMatches()));
  });
}

function setDataStatus(text, status = "loading") {
  els.dataStatus.classList.remove("ready", "warning", "error");
  if (["ready", "warning", "error"].includes(status)) els.dataStatus.classList.add(status);
  els.dataStatus.querySelector("span:last-child").textContent = text;
}

function updateDataMeta() {
  const refreshed = formatRefreshTime(state.data?.meta?.refreshedAt);
  const refreshText = refreshed ? `　刷新于 ${refreshed}` : "";
  const version = dataVersionLabel();
  const versionText = version ? `　数据版本 ${version}` : "";
  els.meta.textContent = `${state.data.meta.source}　${fmtInt.format(state.data.meta.generatedRows)} records${versionText}${refreshText}`;
}

function applyDataPayload(payload) {
  state.data = {
    ...payload,
    meta: { ...payload.meta },
    records: payload.records.map((record, index) => normalizeRecord(record, index)),
  };
  state.voyageCache.clear();
  state.transferCallSelections.clear();
  indexRecords();
  indexTransferEvents();
  els.laneList.innerHTML = optionList(state.data.lanes);
  els.vvdList.innerHTML = optionList(state.data.vvds);
  updateDataMeta();
}

async function refreshBookingData({ manual = false } = {}) {
  if (state.refreshingData) return;
  const refreshApi = window.TSBookingRefresh;
  if (!refreshApi) {
    setDataStatus("刷新组件不可用", "error");
    return;
  }

  state.refreshingData = true;
  els.refreshBookingButton.disabled = true;
  els.refreshBookingButton.setAttribute("aria-busy", "true");
  setDataStatus(manual ? "正在手动刷新…" : "正在检查今日数据…");

  try {
    if (!manual) {
      const existing = await refreshApi.readCachedSnapshot();
      if (existing?.dateKey === refreshApi.singaporeDateKey()) {
        applyDataPayload(refreshApi.mergeSnapshot(state.basePayload, existing));
        refreshVvdList();
        render();
        const version = dataVersionLabel();
        setDataStatus(
          existing.fallbackSource && version
            ? `今日已检查，使用最近可用版 ${version}`
            : `今日已更新 ${formatRefreshTime(existing.refreshedAt)}`,
          "ready",
        );
        return;
      }
    }

    const snapshot = await refreshApi.downloadLatestSnapshot((message) => {
      setDataStatus(message);
    });
    applyDataPayload(refreshApi.mergeSnapshot(state.basePayload, snapshot));
    refreshVvdList();
    render();
    const version = dataVersionLabel();
    setDataStatus(
      snapshot.fallbackSource && version
        ? `最新版不可用，已载入最近可用版 ${version}`
        : `${manual ? "手动刷新完成" : "今日数据已更新"} ${formatRefreshTime(snapshot.refreshedAt)}`,
      "ready",
    );
  } catch (error) {
    console.error("Daily Booking refresh failed", error);
    setDataStatus("GitHub 更新暂不可用，当前数据可继续使用", "warning");
    els.dataStatus.title = error?.message || "无法从 GitHub 刷新数据";
  } finally {
    state.refreshingData = false;
    els.refreshBookingButton.disabled = false;
    els.refreshBookingButton.removeAttribute("aria-busy");
  }
}

async function init() {
  initializeTransferMultiSelects();
  let payload = window.__TS_BOOKING_DATA__ || null;
  const loadSource = payload ? "script" : "json";
  if (!payload) {
    const response = await fetch("data/booking-data.json?v=8");
    if (!response.ok) throw new Error("booking-data.json not found");
    payload = await response.json();
  }
  state.basePayload = payload;

  const refreshApi = window.TSBookingRefresh;
  const cachedSnapshot = refreshApi ? await refreshApi.readCachedSnapshot() : null;
  const initialPayload = cachedSnapshot
    ? refreshApi.mergeSnapshot(state.basePayload, cachedSnapshot)
    : state.basePayload;
  applyDataPayload(initialPayload);

  els.dataStatus.dataset.source = cachedSnapshot ? "browser-cache" : loadSource;
  wireEvents();
  restoreQuery();
  refreshVvdList();
  renderAllTransferMultiSelects();
  render();

  if (!refreshApi) {
    setDataStatus("静态数据已就绪", "ready");
    return;
  }
  if (cachedSnapshot?.dateKey === refreshApi.singaporeDateKey()) {
    const version = dataVersionLabel();
    setDataStatus(
      cachedSnapshot.fallbackSource && version
        ? `今日已检查，使用最近可用版 ${version}`
        : `今日已更新 ${formatRefreshTime(cachedSnapshot.refreshedAt)}`,
      "ready",
    );
    return;
  }

  setDataStatus(cachedSnapshot ? "已载入上次数据，正在更新…" : "数据已就绪，正在更新…");
  await refreshBookingData();
}

init().catch((error) => {
  els.meta.textContent = "数据载入失败";
  els.dataStatus.classList.add("error");
  els.dataStatus.querySelector("span:last-child").textContent = "载入失败";
  showEmpty("数据载入失败", error.message);
});
