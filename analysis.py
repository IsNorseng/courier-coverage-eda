"""Rebuild aggregated courier-coverage results from the original Excel extract."""

import argparse
import csv
import json
import re
import unicodedata
from collections import Counter, defaultdict
from itertools import combinations
from pathlib import Path
from string import ascii_uppercase

from openpyxl import load_workbook


SHEET = "ДатаФрейм"
COLUMNS = ("POSTAL CODE", "ESTADO", "COLONIA", "КС")
STATE_ALIASES = {
    "Aguas Caclientes": "Aguascalientes",
    "Baja California Norte": "Baja California",
    "CDMX": "Ciudad de México",
    "Ciudad de Mexico": "Ciudad de México",
    "Coahuila de Zaragoza": "Coahuila",
    "Estado de Mexico (Metropolitana)": "Estado de México",
    "Estado de Mexico (Toluca)": "Estado de México",
    "México": "Estado de México",
    "Michoacan": "Michoacán",
    "Michoacán de Ocampo": "Michoacán",
    "NuevoLeon": "Nuevo León",
    "Queretaro": "Querétaro",
    "San Luis Potosi": "San Luis Potosí",
    "Veracruz de Ignacio de la Llave": "Veracruz",
    "Yucatan": "Yucatán",
}


def clean_text(value):
    return unicodedata.normalize("NFC", " ".join(str(value or "").split()))


def postal_code(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if not 0 <= value <= 99999 or int(value) != value:
            return None
        value = str(int(value))
    digits = clean_text(value)
    return digits.zfill(5) if re.fullmatch(r"[0-9]{1,5}", digits) else None


def extract(source):
    counts = Counter()
    labels = Counter()
    aliases = Counter()
    named = defaultdict(set)
    missing = []
    postal_all = defaultdict(set)
    carrier_labels = {}
    book = load_workbook(source, read_only=True, data_only=True)
    try:
        rows = book[SHEET].iter_rows(values_only=True)
        header = next(rows, ())
        if not all(name in header for name in COLUMNS):
            raise ValueError(f"Ожидаемые столбцы на листе {SHEET!r}: {COLUMNS}.")
        positions = tuple(header.index(name) for name in COLUMNS)
        for row in rows:
            if not any(value is not None and value != "" for value in row):
                continue
            counts["rows"] += 1
            raw_postal, raw_state, colony, raw_carrier = (row[i] for i in positions)
            label = clean_text(raw_state)
            labels[label] += 1
            counts["empty_colony_rows"] += not clean_text(colony)
            code = postal_code(raw_postal)
            carrier_key = clean_text(raw_carrier)
            if not code or not carrier_key:
                counts["invalid_rows"] += 1
                continue
            if carrier_key not in carrier_labels:
                if len(carrier_labels) == len(ascii_uppercase):
                    raise ValueError("В книге более 26 служб. Для них нужно расширить набор условных обозначений.")
                letter = ascii_uppercase[len(carrier_labels)]
                carrier_labels[carrier_key] = f"Служба {letter}"
            carrier = carrier_labels[carrier_key]
            counts["numeric_postal_rows"] += isinstance(raw_postal, (int, float))
            counts["zero_padded_postal_rows"] += len(clean_text(raw_postal).removesuffix(".0")) < 5
            postal_all[code].add(carrier)
            if label in ("", "-"):
                missing.append((code, carrier))
                continue
            state = STATE_ALIASES.get(label, label)
            if state != label:
                aliases[(label, state)] += 1
            named[(state, code)].add(carrier)
    finally:
        book.close()

    if not named:
        raise ValueError("На листе нет записей с пригодными почтовым индексом, штатом и КС.")
    states_by_code = defaultdict(set)
    for state, code in named:
        states_by_code[code].add(state)

    for code, carrier in missing:
        candidates = states_by_code.get(code, ())
        if len(candidates) == 1:
            state = next(iter(candidates))
            named[(state, code)].add(carrier)
            counts["inferred_state_rows"] += 1
        else:
            counts["unresolved_state_rows"] += 1
            counts["ambiguous_state_rows"] += len(candidates) > 1

    postal = defaultdict(set)
    for (_, code), carriers in named.items():
        postal[code].update(carriers)
    skipped = counts["invalid_rows"] + counts["unresolved_state_rows"]
    included = counts["rows"] - skipped
    source_summary = {
        "sheet": SHEET,
        "rows": counts["rows"],
        "included_rows": included,
        "skipped_rows": skipped,
        "invalid_rows": counts["invalid_rows"],
        "empty_colony_rows": counts["empty_colony_rows"],
        "raw_state_labels": len(labels),
        "normalized_states": len({state for state, _ in named}),
        "numeric_postal_rows": counts["numeric_postal_rows"],
        "zero_padded_postal_rows": counts["zero_padded_postal_rows"],
        "missing_state_rows": len(missing),
        "inferred_state_rows": counts["inferred_state_rows"],
        "unresolved_state_rows": counts["unresolved_state_rows"],
        "ambiguous_state_rows": counts["ambiguous_state_rows"],
        "duplicate_memberships": included - sum(map(len, named.values())),
        "changed_state_rows": sum(aliases.values()),
        "cross_state_postal_codes": sum(len(states) > 1 for states in states_by_code.values()),
        "state_aliases": [
            {"source": original, "state": state, "rows": n}
            for (original, state), n in sorted(aliases.items())
        ],
    }
    return named, postal, postal_all, source_summary


def carrier_sets(index):
    groups = defaultdict(set)
    for key, carriers in index.items():
        for carrier in carriers:
            groups[carrier].add(key)
    return dict(sorted(groups.items()))


def coverage_summary(index, carrier_count):
    histogram = Counter(map(len, index.values()))
    zones = len(index)
    one = histogram[1]
    return {
        "zones": zones,
        "one_carrier": one,
        "at_least_two": zones - one,
        "at_least_two_share": (zones - one) / zones,
        "all_carriers": histogram[carrier_count],
        "carrier_count": carrier_count,
        "memberships": sum(size * n for size, n in histogram.items()),
    }


def pair_metrics(groups):
    for a, b in combinations(groups, 2):
        first, second = groups[a], groups[b]
        intersection = len(first & second)
        union = len(first | second)
        yield {
            "a": a,
            "b": b,
            "intersection": intersection,
            "union": union,
            "jaccard": intersection / union,
            "b_added_to_a": len(second - first),
            "a_added_to_b": len(first - second),
        }


def analyze(source):
    zones, postal, postal_all, source_summary = extract(source)
    groups = carrier_sets(zones)
    summary = coverage_summary(zones, len(groups))
    multiplicity = Counter(map(len, zones.values()))
    exclusive = Counter(next(iter(names)) for names in zones.values() if len(names) == 1)
    state_zones = defaultdict(dict)
    for (state, code), names in zones.items():
        state_zones[state][code] = names

    states = []
    for state, region in sorted(state_zones.items(), key=lambda item: (-len(item[1]), item[0])):
        stats = coverage_summary(region, len(groups))
        volumes = Counter(name for names in region.values() for name in names)
        states.append({
            "state": state,
            "zones": stats["zones"],
            "at_least_two": stats["at_least_two"],
            "at_least_two_share": stats["at_least_two_share"],
            "one_carrier": stats["one_carrier"],
            "by_carrier": {name: volumes[name] for name in groups},
        })

    postal_groups = carrier_sets(postal)
    postal_all_groups = carrier_sets(postal_all)
    result = {
        "source": source_summary,
        "summary": summary,
        "postal_only": coverage_summary(postal, len(postal_groups)),
        "postal_only_all": coverage_summary(postal_all, len(postal_all_groups)),
        "distribution": [
            {"carriers": size, "zones": multiplicity[size], "share": multiplicity[size] / len(zones)}
            for size in range(1, len(groups) + 1)
        ],
        "carriers": [
            {
                "carrier": name,
                "zones": len(area),
                "share": len(area) / len(zones),
                "exclusive": exclusive[name],
                "states": len({state for state, _ in area}),
            }
            for name, area in sorted(groups.items(), key=lambda item: (-len(item[1]), item[0]))
        ],
        "pairs": list(pair_metrics(groups)),
        "postal_only_pairs": list(pair_metrics(postal_groups)),
        "states": states,
    }
    verify(result)
    return result


def verify(result):
    source = result["source"]
    summary = result["summary"]
    postal_only = result["postal_only"]
    distribution = result["distribution"]
    carriers = result["carriers"]
    states = result["states"]
    public_names = {row["carrier"] for row in carriers}
    assert len(public_names) == summary["carrier_count"]
    assert all(re.fullmatch(r"Служба [A-Z]", name) for name in public_names)
    assert all(set(row["by_carrier"]) == public_names for row in states)
    assert source["rows"] == source["included_rows"] + source["skipped_rows"]
    assert source["missing_state_rows"] == source["inferred_state_rows"] + source["unresolved_state_rows"]
    assert source["included_rows"] - source["duplicate_memberships"] == summary["memberships"]
    assert sum(row["zones"] for row in distribution) == summary["zones"]
    assert sum(row["zones"] for row in carriers) == summary["memberships"]
    assert sum(row["exclusive"] for row in carriers) == summary["one_carrier"]
    assert sum(row["zones"] for row in states) == summary["zones"]
    assert sum(row["at_least_two"] for row in states) == summary["at_least_two"]
    assert summary["zones"] >= postal_only["zones"]
    assert summary["memberships"] >= postal_only["memberships"]
    assert result["postal_only_all"]["zones"] >= postal_only["zones"]
    for pair in result["pairs"] + result["postal_only_pairs"]:
        assert {pair["a"], pair["b"]} <= public_names
        assert pair["union"] == pair["intersection"] + pair["b_added_to_a"] + pair["a_added_to_b"]
        assert 0 <= pair["jaccard"] <= 1


def write_csv(path, records):
    if not records:
        return
    with path.open("w", encoding="utf-8-sig", newline="") as output:
        writer = csv.DictWriter(output, fieldnames=records[0].keys())
        writer.writeheader()
        writer.writerows(records)


def write_outputs(result, output):
    output.mkdir(parents=True, exist_ok=True)
    (output / "results.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2) + "\n", encoding="utf-8"
    )
    for name in ("distribution", "carriers", "pairs", "postal_only_pairs"):
        write_csv(output / f"{name}.csv", result[name])
    states = [
        {**{key: value for key, value in row.items() if key != "by_carrier"}, **row["by_carrier"]}
        for row in result["states"]
    ]
    write_csv(output / "states.csv", states)
    write_csv(output / "state_aliases.csv", result["source"]["state_aliases"])


def update_report(report, result):
    marker = '<script type="application/json" id="coverage-data">'
    html = report.read_text(encoding="utf-8")
    if html.count(marker) != 1:
        raise ValueError(f"В отчёте ожидается один блок данных с id=coverage-data: {report}")
    before, _, remainder = html.partition(marker)
    _, close, after = remainder.partition("</script>")
    if not close:
        raise ValueError(f"Блок данных отчёта не закрыт: {report}")
    payload = json.dumps(result, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")
    report.write_text(f"{before}{marker}\n{payload}\n{close}{after}", encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="Пересчитать показатели покрытия из Excel.")
    parser.add_argument("source", type=Path, help="Путь к исходной книге XLSX.")
    parser.add_argument("--output", type=Path, default=Path(__file__).parent / "data", help="Каталог JSON и CSV.")
    parser.add_argument("--report", type=Path, help="Путь к HTML-отчёту для обновления встроенных данных.")
    args = parser.parse_args()
    result = analyze(args.source)
    write_outputs(result, args.output)
    if args.report is not None:
        update_report(args.report, result)
    print(f"Строк: {result['source']['rows']:,}. Зон: {result['summary']['zones']:,}. Результаты: {args.output}")


if __name__ == "__main__":
    main()
