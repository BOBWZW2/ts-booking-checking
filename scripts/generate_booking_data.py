import json
import math
import os
import re
import sys
import urllib.request
import warnings
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "data"
TMP_DIR = Path(os.environ.get("TEMP", ROOT)) / "codex_ts_booking_checking"
DEFAULT_XLSX = TMP_DIR / "booking_latest.xlsx"
DEFAULT_SCHEDULE_XLSX = TMP_DIR / "schedule_latest.xlsx"
BOOKING_URL = "https://raw.githubusercontent.com/BOBWZW2/data-base/main/booking_latest.xlsx"
SCHEDULE_URL = "https://raw.githubusercontent.com/BOBWZW2/data-base/main/schedule_latest.xlsx"
VOYAGE_PATTERN = re.compile(r"^(.*?)(\d{4})([EWNS])$")
OPPOSITE_DIRECTION = {"W": "E", "E": "W", "S": "N", "N": "S"}
MAX_ROUTE_RESOLUTION_HOURS = 90 * 24


def text(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isnan(value):
        return ""
    value = str(value).strip()
    if value.lower() in {"nan", "none"}:
        return ""
    return " ".join(value.split())


def num(value):
    value = text(value).replace(",", "")
    if not value:
        return None
    try:
        return float(value)
    except ValueError:
        return None


def value_or_zero(value):
    parsed = num(value)
    return parsed if parsed is not None else 0.0


def round_qty(value):
    value = 0 if value is None else value
    rounded = round(float(value), 3)
    if abs(rounded) < 0.0005:
        return 0
    if rounded.is_integer():
        return int(rounded)
    return rounded


def leg_column(index, name):
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(index, "th")
    return f"{index}{suffix} {name}"


def adjusted_40(count_20, count_40, teu):
    if teu is None:
        return count_40
    adjusted = (teu - count_20) / 2
    if adjusted < -0.001:
        return count_40
    return adjusted


def stage(row, prefix, fallback_20=None, fallback_40=None, fallback_teu=None):
    count_20 = num(row.get(f"{prefix} Unit 20ft"))
    count_40 = num(row.get(f"{prefix} Unit 40ft"))
    teu = num(row.get(f"{prefix} TTL Teu"))

    if count_20 is None:
        count_20 = value_or_zero(row.get(fallback_20)) if fallback_20 else 0.0
    if count_40 is None:
        count_40 = value_or_zero(row.get(fallback_40)) if fallback_40 else 0.0
    if teu is None:
        teu = num(row.get(fallback_teu)) if fallback_teu else None
    if teu is None:
        teu = count_20 + 2 * count_40

    count_40 = adjusted_40(count_20, count_40, teu)
    return {
        "t20": round_qty(count_20),
        "t40": round_qty(count_40),
        "teu": round_qty(teu),
    }


def max_stage(left, right):
    return {
        "t20": max(float(left["t20"]), float(right["t20"])),
        "t40": max(float(left["t40"]), float(right["t40"])),
        "teu": max(float(left["teu"]), float(right["teu"])),
    }


def normalized_key(value):
    return text(value).upper()


def time_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    if hasattr(value, "strftime"):
        return value.strftime("%Y-%m-%d %H:%M")
    value = text(value)
    if len(value) >= 16 and value[4:5] == "-" and value[7:8] == "-":
        return value[:16]
    return value


def ensure_source(path, url):
    if path.exists():
        return path
    path.parent.mkdir(parents=True, exist_ok=True)
    urllib.request.urlretrieve(url, path)
    return path


def load_schedule():
    source = ensure_source(DEFAULT_SCHEDULE_XLSX, SCHEDULE_URL)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        df = pd.read_excel(source, sheet_name=0, dtype=object)

    by_leg = defaultdict(set)
    by_vvd_port = defaultdict(set)
    voyage_rows = defaultdict(list)

    def put(mapping, key, value):
        if not key or not value:
            return
        mapping[key].add(value)

    for row_index, row in df.iterrows():
        lane = normalized_key(row.get("LANE"))
        vvd = normalized_key(row.get("VVD"))
        port = normalized_key(row.get("PORT"))
        eta = time_text(row.get("ETA"))
        etb = time_text(row.get("ETB"))
        etd = time_text(row.get("ETD"))
        if not vvd or not port:
            continue
        if lane:
            leg_key = f"{lane}|{vvd}|{port}"
            put(by_leg, leg_key, etd)
            source_row = int(row_index) + 2
            voyage_rows[f"{lane}|{vvd}"].append(
                [port, eta, etb, etd, source_row]
            )
        vvd_port_key = f"{vvd}|{port}"
        put(by_vvd_port, vvd_port_key, etd)

    by_leg_payload = {key: sorted(values) for key, values in sorted(by_leg.items())}
    by_vvd_port_payload = {key: sorted(values) for key, values in sorted(by_vvd_port.items())}

    def schedule_time_sort(value):
        return (not bool(value), value)

    def voyage_sort_key(call):
        return (
            schedule_time_sort(call[1]),
            schedule_time_sort(call[2]),
            schedule_time_sort(call[3]),
            call[4],
        )

    calls_by_leg = defaultdict(list)
    calls_by_vvd_port = defaultdict(list)
    voyages_payload = {}
    for voyage_key, raw_calls in sorted(voyage_rows.items()):
        lane, vvd = voyage_key.split("|", 1)
        voyage_calls = []
        for call_seq, raw_call in enumerate(sorted(raw_calls, key=voyage_sort_key), start=1):
            port, eta, etb, etd, source_row = raw_call
            voyage_calls.append([port, eta, etb, etd, call_seq, source_row])
            calls_by_leg[f"{lane}|{vvd}|{port}"].append(
                [eta, etb, etd, call_seq, source_row]
            )
            calls_by_vvd_port[f"{vvd}|{port}"].append(
                [lane, eta, etb, etd, call_seq, source_row]
            )
        voyages_payload[voyage_key] = voyage_calls

    calls_by_leg_payload = {
        key: sorted(values, key=lambda call: (call[3], call[4]))
        for key, values in sorted(calls_by_leg.items())
    }
    calls_by_vvd_port_payload = {
        key: sorted(values, key=lambda call: (call[0], call[4], call[5]))
        for key, values in sorted(calls_by_vvd_port.items())
    }
    exact_conflicts = sum(max(0, len(values) - 1) for values in by_leg_payload.values())
    vvd_port_conflicts = sum(max(0, len(values) - 1) for values in by_vvd_port_payload.values())
    schedule_usable_rows = sum(len(values) for values in voyage_rows.values())
    schedule_multi_call_keys = sum(
        1 for values in calls_by_leg_payload.values() if len(values) > 1
    )
    schedule_max_candidates = max(
        (len(values) for values in calls_by_leg_payload.values()), default=0
    )

    return {
        "byLeg": by_leg_payload,
        "byVvdPort": by_vvd_port_payload,
        "callsByLeg": calls_by_leg_payload,
        "callsByVvdPort": calls_by_vvd_port_payload,
        "voyages": voyages_payload,
        "meta": {
            "scheduleSource": "BOBWZW2/data-base schedule_latest.xlsx",
            "scheduleRows": int(len(df)),
            "scheduleUsableRows": schedule_usable_rows,
            "scheduleSize": source.stat().st_size,
            "scheduleLegKeys": len(by_leg_payload),
            "scheduleVvdPortKeys": len(by_vvd_port_payload),
            "scheduleLegConflicts": exact_conflicts,
            "scheduleVvdPortConflicts": vvd_port_conflicts,
            "scheduleVoyages": len(voyages_payload),
            "scheduleMultiCallKeys": schedule_multi_call_keys,
            "scheduleMaxCandidates": schedule_max_candidates,
        },
    }


def voyage_identity(value):
    match = VOYAGE_PATTERN.match(normalized_key(value))
    if not match:
        return None
    return {
        "vessel": match.group(1),
        "cycle": match.group(2),
        "direction": match.group(3),
    }


def schedule_timestamp(value):
    value = text(value)
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace(" ", "T")).timestamp()
    except ValueError:
        return None


def build_route_overrides(records, schedule):
    """Resolve Booking VVD endpoints onto the actual chronological schedule VVD.

    Booking legs often keep the loading-direction VVD for the whole cargo leg. The
    discharge port can therefore belong to the return direction, for example
    XPBL2625W / CNNGB -> PKKHI where PKKHI is called by XPBL2625E. We keep the
    Booking VVD unchanged and only override the affected schedule endpoint.
    """

    route_counts = Counter()
    for record in records:
        for pol, pod, lane, vvd in record[12]:
            pol = normalized_key(pol)
            pod = normalized_key(pod)
            lane = normalized_key(lane)
            vvd = normalized_key(vvd)
            if pol and pod and lane and vvd:
                route_counts[f"{lane}|{vvd}|{pol}|{pod}"] += 1

    calls_by_voyage_port = defaultdict(list)
    calls_by_lane_family_port = defaultdict(list)
    calls_by_family = defaultdict(list)
    for voyage_key, raw_calls in schedule["voyages"].items():
        lane, vvd = voyage_key.split("|", 1)
        identity = voyage_identity(vvd)
        for raw_call in raw_calls:
            port, eta, etb, etd, call_seq, source_row = raw_call
            call = {
                "lane": lane,
                "vvd": vvd,
                "port": normalized_key(port),
                "eta": eta,
                "etb": etb,
                "etd": etd,
                "callSeq": call_seq,
                "sourceRow": source_row,
            }
            calls_by_voyage_port[(lane, vvd, call["port"])].append(call)
            if identity:
                calls_by_lane_family_port[
                    (lane, identity["vessel"], call["port"])
                ].append(call)
                calls_by_family[identity["vessel"]].append(call)

    def departure_time(call):
        return schedule_timestamp(call["etd"] or call["etb"] or call["eta"])

    def arrival_time(call):
        return schedule_timestamp(call["etb"] or call["eta"] or call["etd"])

    def positive_pairs(departures, arrivals):
        pairs = []
        for departure in departures:
            start = departure_time(departure)
            if start is None:
                continue
            for arrival in arrivals:
                end = arrival_time(arrival)
                if end is None:
                    continue
                gap_hours = (end - start) / 3600
                if 0 <= gap_hours <= MAX_ROUTE_RESOLUTION_HOURS:
                    pairs.append((gap_hours, departure, arrival))
        return sorted(
            pairs,
            key=pair_sort_key,
        )

    def pair_sort_key(item):
        return (
            item[0],
            item[2]["etb"] or item[2]["eta"] or item[2]["etd"],
            item[2]["lane"],
            item[2]["vvd"],
            item[2]["callSeq"],
        )

    def next_voyage_pairs(departures, family_calls, booking_vvd, pod):
        pairs = []
        for departure in departures:
            start = departure_time(departure)
            if start is None:
                continue
            future_calls = []
            for call in family_calls:
                if call["vvd"] == booking_vvd:
                    continue
                end = arrival_time(call)
                if end is None:
                    continue
                gap_hours = (end - start) / 3600
                if 0 <= gap_hours <= MAX_ROUTE_RESOLUTION_HOURS:
                    future_calls.append((gap_hours, call))
            if not future_calls:
                continue
            _, next_call = min(
                future_calls,
                key=lambda item: (
                    item[0],
                    item[1]["etb"] or item[1]["eta"] or item[1]["etd"],
                    item[1]["lane"],
                    item[1]["vvd"],
                    item[1]["callSeq"],
                ),
            )
            next_arrivals = [
                call
                for call in family_calls
                if call["lane"] == next_call["lane"]
                and call["vvd"] == next_call["vvd"]
                and call["port"] == pod
            ]
            pairs.extend(positive_pairs([departure], next_arrivals))
        return sorted(pairs, key=pair_sort_key)

    def has_ordered_exact_path(departures, arrivals):
        return any(
            departure["vvd"] == arrival["vvd"]
            and int(departure["callSeq"]) < int(arrival["callSeq"])
            for departure in departures
            for arrival in arrivals
        )

    def override_call(call):
        return [
            call["vvd"],
            call["eta"],
            call["etb"],
            call["etd"],
            call["callSeq"],
            call["sourceRow"],
            call["lane"],
        ]

    overrides = {}
    direction_counts = Counter()
    corrected_booking_legs = 0
    unresolved_routes = 0
    unresolved_booking_legs = 0

    for route_key, booking_count in sorted(route_counts.items()):
        lane, booking_vvd, pol, pod = route_key.split("|", 3)
        identity = voyage_identity(booking_vvd)
        exact_pol = calls_by_voyage_port.get((lane, booking_vvd, pol), [])
        exact_pod = calls_by_voyage_port.get((lane, booking_vvd, pod), [])

        # Port order is authoritative for an exact Booking VVD. If the source
        # timestamps themselves overlap or invert, keep that real voyage and let
        # the UI flag the timing anomaly instead of silently changing direction.
        if has_ordered_exact_path(exact_pol, exact_pod):
            continue

        # A reliable correction needs the Booking VVD's POL call as the time
        # anchor. Without it, the source does not tell us which cycle loaded the
        # cargo, so keeping "schedule unavailable" is safer than guessing.
        if not identity or not exact_pol:
            unresolved_routes += 1
            unresolved_booking_legs += booking_count
            continue

        expected_direction = OPPOSITE_DIRECTION.get(identity["direction"])
        candidate_pod = []
        for call in calls_by_lane_family_port.get((lane, identity["vessel"], pod), []):
            candidate_identity = voyage_identity(call["vvd"])
            if (
                candidate_identity
                and candidate_identity["direction"] == expected_direction
                and call["vvd"] != booking_vvd
            ):
                candidate_pod.append(call)

        pairs = positive_pairs(exact_pol, candidate_pod)
        if not pairs:
            # A vessel can change service/lane and voyage-direction convention
            # before it reaches the Booking POD. In that case, inspect only the
            # same physical vessel's immediately next scheduled voyage, rather
            # than jumping to any later cycle that happens to call the POD.
            # This resolves rotations such as
            # CCT/CUYP2617N -> REX/CUYP2619W without guessing by voyage letters.
            pairs = next_voyage_pairs(
                exact_pol,
                calls_by_family.get(identity["vessel"], []),
                booking_vvd,
                pod,
            )
        if not pairs:
            unresolved_routes += 1
            unresolved_booking_legs += booking_count
            continue

        gap_hours, departure_call, arrival_call = pairs[0]
        selected_vvd = arrival_call["vvd"]
        selected_lane = arrival_call["lane"]
        selected_arrivals = []
        seen_source_rows = set()
        for _, _, candidate in pairs:
            if candidate["lane"] != selected_lane or candidate["vvd"] != selected_vvd:
                continue
            if candidate["sourceRow"] in seen_source_rows:
                continue
            seen_source_rows.add(candidate["sourceRow"])
            selected_arrivals.append(candidate)
        overrides[route_key] = [
            [],
            [override_call(call) for call in selected_arrivals],
        ]
        direction_counts[
            f"{identity['direction']}>{voyage_identity(arrival_call['vvd'])['direction']}"
        ] += 1
        corrected_booking_legs += booking_count

    return {
        "overrides": overrides,
        "meta": {
            "scheduleRouteOverrideCount": len(overrides),
            "scheduleCorrectedBookingLegs": corrected_booking_legs,
            "scheduleUnresolvedRouteCount": unresolved_routes,
            "scheduleUnresolvedBookingLegs": unresolved_booking_legs,
            "scheduleRouteResolutionMaxHours": MAX_ROUTE_RESOLUTION_HOURS,
            "scheduleRouteOverrideDirections": dict(sorted(direction_counts.items())),
        },
    }


def main():
    source = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_XLSX
    source = ensure_source(source, BOOKING_URL)
    schedule = load_schedule()

    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        df = pd.read_excel(source, sheet_name=0, dtype=object)

    records = []
    lane_vvd_counts = defaultdict(Counter)
    vvd_lane_counts = defaultdict(Counter)
    all_lanes = Counter()
    all_vvds = Counter()
    all_ports = Counter()

    for index, row in df.iterrows():
        booking = stage(
            row,
            "Booking",
            fallback_20="20ft",
            fallback_40="40ft",
            fallback_teu="Booking",
        )
        op = stage(row, "OP")
        vl = stage(row, "VL")
        op = max_stage(op, vl)

        weight_kg = value_or_zero(row.get("Container Weight"))
        booking_weight_ton = weight_kg / 1000 + float(booking["teu"]) * 2

        legs = []
        for leg_index in range(1, 6):
            pol = text(row.get(leg_column(leg_index, "POL")))
            pod = text(row.get(leg_column(leg_index, "POD")))
            lane = text(row.get(leg_column(leg_index, "Lane")))
            vvd = text(row.get(leg_column(leg_index, "VVD")))
            if not (pol or pod or lane or vvd):
                continue
            legs.append({"pol": pol, "pod": pod, "lane": lane, "vvd": vvd})

            if lane:
                all_lanes[lane] += 1
            if vvd:
                all_vvds[vvd] += 1
            if lane and vvd:
                lane_vvd_counts[lane][vvd] += 1
                vvd_lane_counts[vvd][lane] += 1
            if pol:
                all_ports[pol] += 1
            if pod:
                all_ports[pod] += 1

        if not legs:
            continue

        sul_value = text(row.get("SUL YN")).upper()
        record = [
            text(row.get("COC/SOC")).upper(),
            "SUL" if sul_value == "Y" else "CUL",
            booking["t20"],
            booking["t40"],
            booking["teu"],
            op["t20"],
            op["t40"],
            op["teu"],
            vl["t20"],
            vl["t40"],
            vl["teu"],
            round_qty(booking_weight_ton),
            [[leg["pol"], leg["pod"], leg["lane"], leg["vvd"]] for leg in legs],
            text(row.get("BL No.")),
            text(row.get("CUL CODE")),
            text(row.get("POR")),
            text(row.get("POL")),
            text(row.get("POD")),
            text(row.get("DEL")),
        ]
        records.append(record)

    def sorted_keys(counter):
        return [key for key, _ in sorted(counter.items(), key=lambda item: (-item[1], item[0]))]

    route_overrides = build_route_overrides(records, schedule)

    payload = {
        "meta": {
            "source": "BOBWZW2/data-base booking_latest.xlsx",
            "schemaVersion": 5,
            "generatedRows": len(records),
            "xlsxSize": source.stat().st_size,
            **schedule["meta"],
            **route_overrides["meta"],
        },
        "lanes": sorted_keys(all_lanes),
        "vvds": sorted_keys(all_vvds),
        "ports": sorted_keys(all_ports),
        "laneVvds": {
            lane: sorted_keys(counter) for lane, counter in sorted(lane_vvd_counts.items())
        },
        "vvdLanes": {
            vvd: [
                {"lane": lane, "count": count}
                for lane, count in sorted(counter.items(), key=lambda item: (-item[1], item[0]))
            ]
            for vvd, counter in sorted(vvd_lane_counts.items())
        },
        "schedule": {
            "byLeg": schedule["byLeg"],
            "byVvdPort": schedule["byVvdPort"],
            "callsByLeg": schedule["callsByLeg"],
            "callsByVvdPort": schedule["callsByVvdPort"],
            "voyages": schedule["voyages"],
            "routeOverrides": route_overrides["overrides"],
        },
        "records": records,
    }

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    payload_text = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    output = DATA_DIR / "booking-data.json"
    output.write_text(payload_text, encoding="utf-8")
    script_output = DATA_DIR / "booking-data.js"
    script_output.write_text(
        f"window.__TS_BOOKING_DATA__={payload_text};\n",
        encoding="utf-8",
    )
    print(f"Wrote {output} and {script_output} with {len(records):,} records")


if __name__ == "__main__":
    main()
