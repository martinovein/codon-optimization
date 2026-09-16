#!/usr/bin/env python3
"""Codon-optimize amino acid sequences for Saccharomyces cerevisiae.

Usage:
    python codon_optimize_yeast.py JAK-sample-opt-codon-EXAMPLE.xlsx
    python codon_optimize_yeast.py input.xlsx output.xlsx

The script writes two sheets: wide_format and long_format.
"""

from __future__ import annotations

import argparse
import math
import random
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable

from openpyxl import Workbook, load_workbook

try:
    import pandas as pd
except Exception as exc:  # pandas can be unavailable or ABI-broken.
    pd = None
    PANDAS_IMPORT_ERROR = exc
else:
    PANDAS_IMPORT_ERROR = None

try:
    from Bio.Restriction import AllEnzymes
    from Bio.Seq import Seq
except Exception as exc:
    raise SystemExit(
        "Biopython is required for translation and restriction enzyme handling. "
        "Install it with: python -m pip install biopython"
    ) from exc


DEFAULT_INPUT = "JAK-sample-opt-codon-EXAMPLE.xlsx"
DEFAULT_OUTPUT = "JAK-sample-opt-codon-optimized-output.xlsx"
TARGET_TAXID = "4932"
VARIANTS_PER_ROW = 3
MIN_PAIRWISE_DISTANCE = 3
MAX_ATTEMPTS_PER_VARIANT = 5000

REQUIRED_CANONICAL_COLUMNS = {
    "number": ("Number",),
    "name": ("Name",),
    "flank5": ("Fixed 5' DNA flank",),
    "sequence": ("Sequence (AA, DNA, RNA)", "Sequence (AA, DNA, or RNA)"),
    "flank3": ("Fixed 3' DNA flank",),
    "organism": ("Organism of Expression",),
    "sequence_type": ("Sequence Type (required)",),
    "sites_to_avoid": ("Sites to Avoid", "RE Sites to Avoid  (required)", "RE Sites to Avoid"),
}

VALID_AA = set("ACDEFGHIKLMNPQRSTVWY*")

# Saccharomyces cerevisiae S288C codon usage frequencies, expressed as
# relative weights within each amino-acid synonym group.
YEAST_CODON_USAGE = {
    "F": {"TTT": 26.1, "TTC": 18.4},
    "L": {"TTA": 26.2, "TTG": 27.2, "CTT": 12.3, "CTC": 5.4, "CTA": 13.4, "CTG": 10.5},
    "I": {"ATT": 30.1, "ATC": 17.2, "ATA": 17.8},
    "M": {"ATG": 20.9},
    "V": {"GTT": 22.1, "GTC": 11.8, "GTA": 11.8, "GTG": 10.8},
    "S": {"TCT": 23.5, "TCC": 14.2, "TCA": 18.7, "TCG": 8.6, "AGT": 14.2, "AGC": 9.8},
    "P": {"CCT": 13.5, "CCC": 6.8, "CCA": 18.2, "CCG": 5.3},
    "T": {"ACT": 20.3, "ACC": 12.7, "ACA": 17.8, "ACG": 8.0},
    "A": {"GCT": 21.2, "GCC": 12.6, "GCA": 16.2, "GCG": 6.2},
    "Y": {"TAT": 18.8, "TAC": 14.8},
    "H": {"CAT": 13.6, "CAC": 7.8},
    "Q": {"CAA": 27.3, "CAG": 12.1},
    "N": {"AAT": 35.7, "AAC": 24.8},
    "K": {"AAA": 42.0, "AAG": 30.8},
    "D": {"GAT": 37.6, "GAC": 20.2},
    "E": {"GAA": 45.6, "GAG": 19.2},
    "C": {"TGT": 8.1, "TGC": 4.8},
    "W": {"TGG": 10.4},
    "R": {"CGT": 6.4, "CGC": 2.6, "CGA": 3.0, "CGG": 1.7, "AGA": 21.3, "AGG": 9.2},
    "G": {"GGT": 23.9, "GGC": 9.8, "GGA": 10.9, "GGG": 6.0},
    "*": {"TAA": 1.0, "TAG": 0.5, "TGA": 0.7},
}


def normalize_header(value: Any) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"[^a-z0-9]+", "", text.lower())


def clean_sequence(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return re.sub(r"\s+", "", str(value)).upper()


def value_to_text(value: Any) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ""
    return str(value)


def dna_for_search(value: Any) -> str:
    return clean_sequence(value).replace("U", "T")


def build_column_map(headers: list[str]) -> dict[str, str]:
    normalized_to_actual = {normalize_header(header): header for header in headers}
    column_map: dict[str, str] = {}
    missing: list[str] = []

    for canonical, aliases in REQUIRED_CANONICAL_COLUMNS.items():
        actual = None
        for alias in aliases:
            actual = normalized_to_actual.get(normalize_header(alias))
            if actual is not None:
                break
        if actual is None:
            missing.append(" / ".join(aliases))
        else:
            column_map[canonical] = actual

    if missing:
        raise ValueError("Missing required column(s): " + "; ".join(missing))
    return column_map


def read_excel_rows(path: Path) -> tuple[list[dict[str, Any]], list[str]]:
    if pd is not None:
        df = pd.read_excel(path, sheet_name=0, dtype=object)
        headers = list(df.columns)
        return df.to_dict(orient="records"), headers

    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
    header_text = [value_to_text(header) for header in headers]
    rows = []
    for values in worksheet.iter_rows(min_row=2, values_only=True):
        rows.append({header_text[index]: values[index] for index in range(len(header_text))})
    return rows, header_text


def write_excel(path: Path, wide_rows: list[dict[str, Any]], long_rows: list[dict[str, Any]]) -> None:
    wide_headers = list(wide_rows[0].keys()) if wide_rows else []
    long_headers = [
        "Number",
        "Name",
        "Variant",
        "Fixed 5' DNA flank",
        "Optimized DNA",
        "Fixed 3' DNA flank",
        "Full construct",
        "Status",
        "Warnings",
    ]

    if pd is not None:
        with pd.ExcelWriter(path, engine="openpyxl") as writer:
            pd.DataFrame(wide_rows, columns=wide_headers).to_excel(writer, sheet_name="wide_format", index=False)
            pd.DataFrame(long_rows, columns=long_headers).to_excel(writer, sheet_name="long_format", index=False)
        return

    workbook = Workbook()
    ws_wide = workbook.active
    ws_wide.title = "wide_format"
    ws_wide.append(wide_headers)
    for row in wide_rows:
        ws_wide.append([row.get(header, "") for header in wide_headers])

    ws_long = workbook.create_sheet("long_format")
    ws_long.append(long_headers)
    for row in long_rows:
        ws_long.append([row.get(header, "") for header in long_headers])

    workbook.save(path)


def parse_enzyme_names(value: Any) -> list[str]:
    text = value_to_text(value)
    if not text:
        return []
    tokens = [token for token in re.split(r"[,;\s]+", text.strip()) if token]
    return tokens


def enzyme_lookup() -> dict[str, Any]:
    return {enzyme.__name__.lower(): enzyme for enzyme in AllEnzymes}


ENZYME_LOOKUP = enzyme_lookup()


def resolve_enzymes(names: Iterable[str]) -> tuple[list[Any], list[str]]:
    enzymes = []
    warnings = []
    seen = set()
    for raw_name in names:
        key = raw_name.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        enzyme = ENZYME_LOOKUP.get(key)
        if enzyme is None:
            warnings.append(f"Unrecognized restriction enzyme skipped: {raw_name}")
        else:
            enzymes.append(enzyme)
    return enzymes, warnings


def enzyme_site_counts(sequence: str, enzymes: Iterable[Any]) -> Counter:
    counts: Counter = Counter()
    if not sequence:
        return counts
    seq = Seq(dna_for_search(sequence))
    for enzyme in enzymes:
        try:
            counts[enzyme.__name__] = len(enzyme.search(seq, linear=True))
        except TypeError:
            counts[enzyme.__name__] = len(enzyme.search(seq))
    return counts


def introduced_sites(
    coding_dna: str,
    flank5: str,
    flank3: str,
    enzymes: Iterable[Any],
) -> list[str]:
    enzymes = list(enzymes)
    coding_counts = enzyme_site_counts(coding_dna, enzymes)
    coding_hits = sorted(name for name, count in coding_counts.items() if count > 0)
    if coding_hits:
        return [f"Forbidden site inside optimized coding DNA: {', '.join(coding_hits)}"]

    baseline_counts = enzyme_site_counts(flank5, enzymes) + enzyme_site_counts(flank3, enzymes)
    full_construct = f"{flank5}{coding_dna}{flank3}"
    full_counts = enzyme_site_counts(full_construct, enzymes)

    introduced = []
    for enzyme in enzymes:
        name = enzyme.__name__
        extra_count = full_counts.get(name, 0) - baseline_counts.get(name, 0)
        if extra_count > 0:
            introduced.append(f"{name} ({extra_count} new)")
    if introduced:
        return ["Forbidden site newly introduced in full construct: " + ", ".join(introduced)]
    return []


def random_codon_for_amino_acid(amino_acid: str) -> str:
    choices = YEAST_CODON_USAGE[amino_acid]
    codons = list(choices.keys())
    weights = list(choices.values())
    return random.choices(codons, weights=weights, k=1)[0]


def optimize_once(amino_acids: str) -> str:
    return "".join(random_codon_for_amino_acid(amino_acid) for amino_acid in amino_acids)


def translate_dna(dna_sequence: str) -> str:
    return str(Seq(dna_sequence).translate(table=1, to_stop=False))


def hamming_distance(left: str, right: str) -> int:
    if len(left) != len(right):
        return max(len(left), len(right))
    return sum(a != b for a, b in zip(left, right))


def is_distinct_enough(candidate: str, existing_variants: Iterable[str]) -> bool:
    return all(hamming_distance(candidate, existing) >= MIN_PAIRWISE_DISTANCE for existing in existing_variants)


def validate_amino_acids(amino_acids: str) -> list[str]:
    if not amino_acids:
        return ["Amino acid sequence is empty"]
    invalid = sorted(set(amino_acids) - VALID_AA)
    if invalid:
        return ["Invalid amino acid character(s): " + ", ".join(invalid)]
    return []


def generate_variants(
    amino_acids: str,
    flank5: str,
    flank3: str,
    enzymes: list[Any],
) -> tuple[list[str], list[str]]:
    variants: list[str] = []
    failure_reasons: Counter = Counter()

    for _variant_index in range(VARIANTS_PER_ROW):
        accepted = None
        for _attempt in range(MAX_ATTEMPTS_PER_VARIANT):
            candidate = optimize_once(amino_acids)

            translated = translate_dna(candidate)
            if translated != amino_acids:
                failure_reasons["translation mismatch"] += 1
                continue

            site_failures = introduced_sites(candidate, flank5, flank3, enzymes)
            if site_failures:
                failure_reasons[site_failures[0]] += 1
                continue

            if not is_distinct_enough(candidate, variants):
                failure_reasons[f"variant differs by fewer than {MIN_PAIRWISE_DISTANCE} nt"] += 1
                continue

            accepted = candidate
            break

        if accepted is None:
            break
        variants.append(accepted)

    messages = []
    if len(variants) < VARIANTS_PER_ROW:
        if failure_reasons:
            common = "; ".join(f"{reason} ({count} attempts)" for reason, count in failure_reasons.most_common(3))
            messages.append(
                f"Generated {len(variants)} of {VARIANTS_PER_ROW} variants after retry limit; common rejections: {common}"
            )
        else:
            messages.append(f"Generated {len(variants)} of {VARIANTS_PER_ROW} variants after retry limit")
    return variants, messages


def organism_warning(value: Any) -> str | None:
    text = value_to_text(value).strip()
    if not text:
        return "Organism of Expression is blank; using Saccharomyces cerevisiae taxid 4932"
    if text.endswith(".0"):
        text = text[:-2]
    if text != TARGET_TAXID:
        return f"Organism of Expression is {value_to_text(value)}; using Saccharomyces cerevisiae taxid 4932"
    return None


def process_row(row: dict[str, Any], headers: list[str], column_map: dict[str, str]) -> tuple[dict[str, Any], list[dict[str, Any]], bool]:
    output_row = {header: row.get(header, "") for header in headers}
    warnings = []
    status = "Success"

    flank5 = value_to_text(row.get(column_map["flank5"], ""))
    flank3 = value_to_text(row.get(column_map["flank3"], ""))
    amino_acids = clean_sequence(row.get(column_map["sequence"], ""))
    name = value_to_text(row.get(column_map["name"], ""))

    org_warning = organism_warning(row.get(column_map["organism"], ""))
    if org_warning:
        warnings.append(org_warning)

    enzymes, enzyme_warnings = resolve_enzymes(parse_enzyme_names(row.get(column_map["sites_to_avoid"], "")))
    warnings.extend(enzyme_warnings)

    validation_errors = validate_amino_acids(amino_acids)
    variants: list[str] = []
    failure_messages: list[str] = []
    if validation_errors:
        status = "Failed: " + "; ".join(validation_errors)
    else:
        variants, failure_messages = generate_variants(amino_acids, flank5, flank3, enzymes)
        if len(variants) != VARIANTS_PER_ROW:
            status = "Failed: " + "; ".join(failure_messages)

    for index in range(VARIANTS_PER_ROW):
        optimized = variants[index] if index < len(variants) else ""
        output_row[f"Optimized DNA {index + 1}"] = optimized
        output_row[f"Full construct {index + 1}"] = f"{flank5}{optimized}{flank3}" if optimized else ""

    output_row["Status"] = status
    output_row["Warnings"] = "; ".join(warnings)

    long_rows = []
    for index in range(VARIANTS_PER_ROW):
        optimized = variants[index] if index < len(variants) else ""
        long_rows.append(
            {
                "Number": row.get(column_map["number"], ""),
                "Name": f"{name}_opt{index + 1}" if name else f"_opt{index + 1}",
                "Variant": index + 1,
                "Fixed 5' DNA flank": flank5,
                "Optimized DNA": optimized,
                "Fixed 3' DNA flank": flank3,
                "Full construct": f"{flank5}{optimized}{flank3}" if optimized else "",
                "Status": status,
                "Warnings": "; ".join(warnings),
            }
        )

    return output_row, long_rows, status == "Success"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input_excel", nargs="?", default=DEFAULT_INPUT)
    parser.add_argument("output_excel", nargs="?", default=DEFAULT_OUTPUT)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)
    input_path = Path(args.input_excel)
    output_path = Path(args.output_excel)

    if not input_path.exists():
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    rows, headers = read_excel_rows(input_path)
    column_map = build_column_map(headers)

    if PANDAS_IMPORT_ERROR is not None:
        print(f"Warning: pandas import failed; using openpyxl fallback: {PANDAS_IMPORT_ERROR}", file=sys.stderr)

    wide_rows = []
    long_rows = []
    successful_rows = 0

    for row in rows:
        wide_row, row_long_rows, success = process_row(row, headers, column_map)
        wide_rows.append(wide_row)
        long_rows.extend(row_long_rows)
        if success:
            successful_rows += 1

    write_excel(output_path, wide_rows, long_rows)

    failed_rows = len(rows) - successful_rows
    print(f"Rows processed: {len(rows)}")
    print(f"Successful rows: {successful_rows}")
    print(f"Failed rows: {failed_rows}")
    print(f"Output file: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
