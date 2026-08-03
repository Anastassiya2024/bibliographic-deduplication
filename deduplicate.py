from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import unicodedata
from collections import defaultdict
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable

import pandas as pd


DEFAULT_FILES = [
    Path("savedrecs_1965.xls"),
    Path("savedrecs1_1000.xls"),
    Path("scopus_export_Aug 3-2026_1bcfb716-8c67-418a-82b1-e696e3343e87.csv"),
    Path("csv-BreastNeop-set-3(1).csv"),
]

OUTPUT_DIR = Path("deduplication")

SOURCE_LABELS = {
    "savedrecs_1965.xls": "Web of Science",
    "savedrecs1_1000.xls": "Web of Science",
    "scopus_export_Aug 3-2026_1bcfb716-8c67-418a-82b1-e696e3343e87.csv": "Scopus",
    "csv-BreastNeop-set-3(1).csv": "PubMed",
}

FIELD_ALIASES = {
    "title": [
        "title", "article title", "document title", "ti",
    ],
    "doi": [
        "doi", "digital object identifier", "di",
    ],
    "pmid": [
        "pmid", "pubmed id", "pubmed_id",
    ],
    "authors": [
        "authors", "author full names", "author(s)", "au",
    ],
    "first_author": [
        "first author", "first_author",
    ],
    "year": [
        "year", "publication year", "publication_year", "py",
    ],
    "journal": [
        "journal/book", "source title", "journal", "publication name", "so",
    ],
    "abstract": [
        "abstract", "ab",
    ],
    "keywords": [
        "author keywords", "index keywords", "keywords", "de", "id",
    ],
    "pmcid": [
        "pmcid",
    ],
    "wos_id": [
        "ut (unique wos id)", "ut", "accession number",
    ],
    "scopus_eid": [
        "eid",
    ],
    "url": [
        "link", "url",
    ],
    "document_type": [
        "document type", "publication type", "publication types", "dt",
    ],
}


def clean_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).lower()


def first_matching_column(columns: Iterable[str], aliases: list[str]) -> str | None:
    normalized = {clean_header(c): c for c in columns}
    for alias in aliases:
        if clean_header(alias) in normalized:
            return normalized[clean_header(alias)]
    return None


def safe_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_doi(value: Any) -> str:
    text = safe_text(value).lower()
    text = re.sub(r"^https?://(?:dx\.)?doi\.org/", "", text)
    text = re.sub(r"^doi:\s*", "", text)
    text = text.strip().strip(" .;,")
    match = re.search(r"10\.\d{4,9}/\S+", text)
    if match:
        return match.group(0).rstrip(" .;,)")
    return ""


def normalize_pmid(value: Any) -> str:
    text = safe_text(value)
    match = re.search(r"\b\d{6,9}\b", text)
    return match.group(0) if match else ""


def normalize_year(value: Any) -> str:
    text = safe_text(value)
    match = re.search(r"\b(19|20)\d{2}\b", text)
    return match.group(0) if match else ""


def normalize_title(value: Any) -> str:
    text = safe_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = text.replace("&", " and ")
    text = re.sub(r"\b(article in press|ahead of print|early view)\b", " ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def normalize_author(value: Any) -> str:
    text = safe_text(value).lower()
    if not text:
        return ""
    # Keep first author surname-like token.
    text = re.split(r"[;,|]", text)[0]
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    tokens = re.findall(r"[a-z]+", text)
    return tokens[0] if tokens else ""


def title_tokens(title_norm: str) -> set[str]:
    stop = {
        "a", "an", "and", "as", "at", "by", "for", "from", "in", "of",
        "on", "or", "the", "to", "with", "among", "patients", "study",
    }
    return {t for t in title_norm.split() if len(t) > 2 and t not in stop}


def jaccard(a: set[str], b: set[str]) -> float:
    if not a or not b:
        return 0.0
    return len(a & b) / len(a | b)


def read_input(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, dtype=str, encoding="utf-8-sig", low_memory=False)
    if suffix in {".xls", ".xlsx"}:
        try:
            return pd.read_excel(path, dtype=str)
        except ImportError as exc:
            raise RuntimeError(
                f"Cannot read {path.name}. Install Excel support with:\n"
                "python -m pip install xlrd openpyxl"
            ) from exc
    raise ValueError(f"Unsupported file type: {path}")


def infer_source(path: Path) -> str:
    if path.name in SOURCE_LABELS:
        return SOURCE_LABELS[path.name]
    lower = path.name.lower()
    if "scopus" in lower:
        return "Scopus"
    if "savedrecs" in lower or "wos" in lower:
        return "Web of Science"
    if "pubmed" in lower or "breastneop" in lower:
        return "PubMed"
    return path.stem


def standardize_file(path: Path) -> list[dict[str, Any]]:
    df = read_input(path)
    source = infer_source(path)
    columns = list(df.columns)

    mapped = {
        field: first_matching_column(columns, aliases)
        for field, aliases in FIELD_ALIASES.items()
    }

    if not mapped["title"]:
        raise ValueError(
            f"No title column recognized in {path.name}. Columns: {columns}"
        )

    rows: list[dict[str, Any]] = []
    for index, raw in df.iterrows():
        def get(field: str) -> str:
            col = mapped.get(field)
            return safe_text(raw.get(col, "")) if col else ""

        title = get("title")
        if not title:
            continue

        authors = get("authors")
        first_author = get("first_author") or normalize_author(authors)

        row = {
            "record_id": "",
            "source_database": source,
            "source_file": path.name,
            "source_row": int(index) + 2,
            "title": title,
            "authors": authors,
            "first_author": first_author,
            "year": normalize_year(get("year")),
            "journal": get("journal"),
            "doi": normalize_doi(get("doi")),
            "pmid": normalize_pmid(get("pmid")),
            "pmcid": get("pmcid"),
            "wos_id": get("wos_id"),
            "scopus_eid": get("scopus_eid"),
            "abstract": get("abstract"),
            "keywords": get("keywords"),
            "document_type": get("document_type"),
            "url": get("url"),
        }
        row["title_normalized"] = normalize_title(title)
        row["first_author_normalized"] = normalize_author(first_author or authors)
        row["record_id"] = hashlib.sha1(
            f"{source}|{path.name}|{index}".encode("utf-8")
        ).hexdigest()[:14]
        rows.append(row)

    return rows


def completeness_score(row: dict[str, Any]) -> tuple[int, int, int]:
    populated = sum(
        bool(safe_text(row.get(field)))
        for field in [
            "doi", "pmid", "pmcid", "wos_id", "scopus_eid", "authors",
            "year", "journal", "abstract", "keywords", "document_type", "url",
        ]
    )
    return (
        populated,
        len(safe_text(row.get("abstract"))),
        len(safe_text(row.get("authors"))),
    )


def merge_sources(keeper: dict[str, Any], duplicate: dict[str, Any]) -> None:
    sources = set(filter(None, safe_text(keeper.get("source_database")).split("; ")))
    sources.add(safe_text(duplicate.get("source_database")))
    keeper["source_database"] = "; ".join(sorted(sources))

    files = set(filter(None, safe_text(keeper.get("source_file")).split("; ")))
    files.add(safe_text(duplicate.get("source_file")))
    keeper["source_file"] = "; ".join(sorted(files))

    for field in [
        "doi", "pmid", "pmcid", "wos_id", "scopus_eid", "authors",
        "first_author", "year", "journal", "abstract", "keywords",
        "document_type", "url",
    ]:
        if not safe_text(keeper.get(field)) and safe_text(duplicate.get(field)):
            keeper[field] = duplicate[field]


def exact_deduplicate(records: list[dict[str, Any]]):
    active: dict[str, dict[str, Any]] = {r["record_id"]: dict(r) for r in records}
    removed: list[dict[str, Any]] = []

    def process_key(key_name: str, key_func):
        groups: dict[str, list[str]] = defaultdict(list)
        for rid, row in active.items():
            key = key_func(row)
            if key:
                groups[key].append(rid)

        for key, ids in groups.items():
            if len(ids) < 2:
                continue
            rows = [active[rid] for rid in ids if rid in active]
            if len(rows) < 2:
                continue
            rows.sort(key=completeness_score, reverse=True)
            keeper = rows[0]
            for dup in rows[1:]:
                if dup["record_id"] not in active:
                    continue
                merge_sources(keeper, dup)
                removed.append({
                    "duplicate_record_id": dup["record_id"],
                    "kept_record_id": keeper["record_id"],
                    "match_stage": key_name,
                    "match_value": key,
                    "duplicate_source": dup["source_database"],
                    "kept_source": keeper["source_database"],
                    "duplicate_title": dup["title"],
                    "kept_title": keeper["title"],
                    "duplicate_doi": dup["doi"],
                    "kept_doi": keeper["doi"],
                    "duplicate_pmid": dup["pmid"],
                    "kept_pmid": keeper["pmid"],
                })
                del active[dup["record_id"]]

    process_key("exact_doi", lambda r: r["doi"])
    process_key("exact_pmid", lambda r: r["pmid"])
    process_key(
        "exact_title_year",
        lambda r: (
            f"{r['title_normalized']}|{r['year']}"
            if r["title_normalized"] and r["year"] else ""
        ),
    )
    process_key(
        "exact_title",
        lambda r: (
            r["title_normalized"]
            if len(r["title_normalized"]) >= 35 else ""
        ),
    )

    return list(active.values()), removed


def build_blocks(records: list[dict[str, Any]]) -> dict[str, list[int]]:
    blocks: dict[str, list[int]] = defaultdict(list)
    for i, row in enumerate(records):
        title = row["title_normalized"]
        tokens = sorted(title_tokens(title))
        year = row["year"] or "unknown"
        author = row["first_author_normalized"] or "unknown"

        if tokens:
            key_tokens = "_".join(tokens[:3])
            blocks[f"{year}|{key_tokens}"].append(i)
            blocks[f"{author}|{key_tokens}"].append(i)

        if len(title) >= 20:
            blocks[f"{year}|prefix|{title[:25]}"].append(i)
    return blocks


def identify_possible_duplicates(
    records: list[dict[str, Any]],
    similarity_threshold: float = 0.92,
) -> list[dict[str, Any]]:
    blocks = build_blocks(records)
    candidate_pairs: set[tuple[int, int]] = set()

    for indices in blocks.values():
        if len(indices) < 2 or len(indices) > 100:
            continue
        for pos, i in enumerate(indices):
            for j in indices[pos + 1:]:
                if i != j:
                    candidate_pairs.add((min(i, j), max(i, j)))

    possible: list[dict[str, Any]] = []
    for i, j in sorted(candidate_pairs):
        a, b = records[i], records[j]

        # Already exact identifiers disagree only if both are present and different.
        if a["doi"] and b["doi"] and a["doi"] != b["doi"]:
            continue
        if a["pmid"] and b["pmid"] and a["pmid"] != b["pmid"]:
            continue

        year_compatible = (
            not a["year"] or not b["year"]
            or abs(int(a["year"]) - int(b["year"])) <= 1
        )
        if not year_compatible:
            continue

        ratio = SequenceMatcher(
            None, a["title_normalized"], b["title_normalized"]
        ).ratio()
        jac = jaccard(
            title_tokens(a["title_normalized"]),
            title_tokens(b["title_normalized"]),
        )
        author_match = (
            bool(a["first_author_normalized"])
            and a["first_author_normalized"] == b["first_author_normalized"]
        )

        if ratio < similarity_threshold and jac < 0.88:
            continue

        if ratio >= 0.985 and (author_match or jac >= 0.95):
            recommendation = "Likely duplicate"
        elif ratio >= 0.96 and (author_match or jac >= 0.90):
            recommendation = "Probable duplicate"
        else:
            recommendation = "Manual review"

        possible.append({
            "pair_id": f"P{len(possible)+1:05d}",
            "record_id_a": a["record_id"],
            "record_id_b": b["record_id"],
            "source_a": a["source_database"],
            "source_b": b["source_database"],
            "title_a": a["title"],
            "title_b": b["title"],
            "year_a": a["year"],
            "year_b": b["year"],
            "first_author_a": a["first_author"],
            "first_author_b": b["first_author"],
            "doi_a": a["doi"],
            "doi_b": b["doi"],
            "pmid_a": a["pmid"],
            "pmid_b": b["pmid"],
            "title_similarity": round(ratio, 4),
            "token_jaccard": round(jac, 4),
            "first_author_match": "Yes" if author_match else "No",
            "recommendation": recommendation,
            "review_decision": "Pending",
            "reviewer_note": "",
        })
    return possible


def write_csv(path: Path, rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge PubMed, Scopus and Web of Science exports and deduplicate them."
    )
    parser.add_argument(
        "files",
        nargs="*",
        type=Path,
        help="Input .csv/.xls/.xlsx files. If omitted, known filenames are used.",
    )
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument(
        "--similarity-threshold",
        type=float,
        default=0.92,
        help="Minimum fuzzy title similarity for possible-duplicate review.",
    )
    args = parser.parse_args()

    files = args.files or [p for p in DEFAULT_FILES if p.exists()]
    if not files:
        raise FileNotFoundError(
            "No input files found. Provide paths explicitly or place exports in the project root."
        )

    all_records: list[dict[str, Any]] = []
    source_counts: dict[str, int] = defaultdict(int)

    for path in files:
        if not path.exists():
            raise FileNotFoundError(path)
        rows = standardize_file(path)
        all_records.extend(rows)
        source_counts[infer_source(path)] += len(rows)
        print(f"Loaded {len(rows)} records from {path.name}")

    unique_records, exact_duplicates = exact_deduplicate(all_records)
    possible_duplicates = identify_possible_duplicates(
        unique_records,
        similarity_threshold=args.similarity_threshold,
    )

    args.output_dir.mkdir(parents=True, exist_ok=True)
    write_csv(args.output_dir / "all_records_standardized.csv", all_records)
    write_csv(args.output_dir / "unique_records_exact_dedup.csv", unique_records)
    write_csv(args.output_dir / "duplicates_exact_removed.csv", exact_duplicates)
    write_csv(args.output_dir / "possible_duplicates_review.csv", possible_duplicates)

    summary = {
        "input_files": [str(p) for p in files],
        "records_by_source": dict(sorted(source_counts.items())),
        "total_records_imported": len(all_records),
        "exact_duplicates_removed": len(exact_duplicates),
        "records_after_exact_deduplication": len(unique_records),
        "possible_duplicate_pairs_for_manual_review": len(possible_duplicates),
        "final_screening_count_before_manual_duplicate_review": len(unique_records),
        "deduplication_rules": [
            "exact normalized DOI",
            "exact PMID",
            "exact normalized title plus publication year",
            "exact normalized title",
            "fuzzy title matching for manual review only",
        ],
    }

    (args.output_dir / "dedup_summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )

    text = [
        "Deduplication summary",
        "=====================",
        *[
            f"{source}: {count}"
            for source, count in sorted(source_counts.items())
        ],
        f"Total imported: {len(all_records)}",
        f"Exact duplicates removed: {len(exact_duplicates)}",
        f"Records after exact deduplication: {len(unique_records)}",
        f"Possible duplicate pairs requiring review: {len(possible_duplicates)}",
        "",
        "Important:",
        "Possible duplicate pairs are not removed automatically.",
        "Review possible_duplicates_review.csv and confirm each decision.",
    ]
    (args.output_dir / "dedup_summary.txt").write_text(
        "\n".join(text) + "\n",
        encoding="utf-8",
    )

    print()
    print("\n".join(text))


if __name__ == "__main__":
    main()
