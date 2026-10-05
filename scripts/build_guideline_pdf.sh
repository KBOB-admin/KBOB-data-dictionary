#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
output="$repo_root/docs/Leitfaden_Strukturvorlage_Housekeeping_und_Anleitung.pdf"

requested_font="${GUIDE_FONT:-Arial}"
resolved_family="$(fc-match -f '%{family}\n' "$requested_font" | head -n 1)"
if [[ "$resolved_family" == Arial* ]]; then
  main_font="$requested_font"
else
  main_font="Liberation Sans"
  echo "Hinweis: Arial ist nicht installiert; für das PDF wird Liberation Sans als metrisch kompatibler Ersatz verwendet." >&2
fi

pandoc \
  "$repo_root/docs/repository-housekeeping.md" \
  "$repo_root/docs/strukturvorlage-schritt-fuer-schritt.md" \
  --from=gfm \
  --pdf-engine=xelatex \
  --metadata lang=de-CH \
  --metadata title="Leitfaden Strukturvorlage: Housekeeping und Schritt-für-Schritt-Anleitung" \
  --metadata author="KBOB Data Dictionary Repository" \
  --toc \
  -V toc-title="Inhaltsverzeichnis" \
  -V mainfont="$main_font" \
  -V sansfont="$main_font" \
  -V monofont="Liberation Mono" \
  -V papersize=a4 \
  -V geometry:margin=25mm \
  -V fontsize=11pt \
  -V linestretch=1.15 \
  -V colorlinks=false \
  -V linkcolor=black \
  -V urlcolor=black \
  -o "$output"

echo "$output"
