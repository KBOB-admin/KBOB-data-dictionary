#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output="$repo_root/docs/Leitfaden_Strukturvorlage_Housekeeping_und_Anleitung.docx"

pandoc \
  "$repo_root/docs/repository-housekeeping.md" \
  "$repo_root/docs/strukturvorlage-schritt-fuer-schritt.md" \
  --from=gfm \
  --metadata lang=de-CH \
  --metadata title="Leitfaden Strukturvorlage: Housekeeping und Schritt-für-Schritt-Anleitung" \
  --metadata author="KBOB Data Dictionary Repository" \
  --toc \
  -o "$output"

work_dir="$(mktemp -d)"
trap 'rm -rf "$work_dir"' EXIT
unzip -q "$output" -d "$work_dir"

# Pandoc uses the Microsoft Office theme fonts Aptos/Aptos Display by default.
# Replacing the theme font names preserves the Word styles while making all
# normal text and headings explicitly use Arial on systems where it is present.
sed -i 's/typeface="Aptos Display"/typeface="Arial"/g; s/typeface="Aptos"/typeface="Arial"/g' \
  "$work_dir/word/theme/theme1.xml"

(
  cd "$work_dir"
  zip -q -r "$output" .
)

echo "$output"
