# Bibliographic Deduplication

A Python tool for merging and deduplicating bibliographic records exported from **PubMed**, **Scopus**, and **Web of Science Core Collection**. The tool harmonizes metadata across databases, removes exact duplicates, identifies potential duplicates using fuzzy matching, and generates standardized outputs for systematic reviews.

## Features

- Harmonizes bibliographic fields across databases
- Exact duplicate detection using DOI and PMID
- Normalized title matching
- Fuzzy title matching for potential duplicates
- Generates a manual review list for uncertain duplicates
- Produces standardized datasets suitable for PRISMA-based systematic reviews

## Supported input formats

- CSV
- XLS
- XLSX

## Supported databases

- PubMed (MEDLINE)
- Scopus
- Web of Science Core Collection

## Output files

- `all_records_standardized.csv`
- `unique_records_exact_dedup.csv`
- `duplicates_exact_removed.csv`
- `possible_duplicates_review.csv`
- `dedup_summary.json`
- `dedup_summary.txt`

## Deduplication workflow

1. Import bibliographic records from multiple databases.
2. Harmonize field names and metadata.
3. Normalize titles, authors, DOI, PMID, and publication year.
4. Remove exact duplicates using:
   - DOI
   - PMID
   - normalized title + publication year
   - normalized title
5. Identify potential duplicates using fuzzy title matching.
6. Generate a manual review file for uncertain duplicate pairs.

## Requirements

- Python 3.10+
- pandas
- openpyxl
- xlrd

## License

MIT License.
