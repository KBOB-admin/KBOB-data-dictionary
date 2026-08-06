from __future__ import annotations

import argparse
import json
import re
import time
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
CACHE_JSON = REPO_ROOT / 'resources' / 'bsdd' / 'ifc4.3-uri-cache.json'
URI_RE = re.compile(r'https://identifier\.buildingsmart\.org/uri/buildingsmart/ifc/4\.3(?:\.0)?/[^\s<>"]+')


def build_cache(source: Path, target: Path = CACHE_JSON) -> dict:
    if not source.exists():
        raise FileNotFoundError(f'bSDD source TTL not found: {source}')
    start = time.time()
    uris = set()
    with source.open('r', encoding='utf-8', errors='ignore') as f:
        for line in f:
            for m in URI_RE.findall(line):
                # normalize obsolete '/ifc/4.3.0/' → '/ifc/4.3/' in extracted URIs only
                norm = m.replace('/ifc/4.3.0/','/ifc/4.3/')
                uris.add(norm)
    payload = {
        'source': str(source),
        'uri_count': len(uris),
        'uris': sorted(uris),
        'generated_at_epoch': time.time(),
        'elapsed_seconds': round(time.time() - start, 3),
    }
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False))
    return payload


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Build the repository-local IFC/bSDD URI cache from a harvested bSDD TTL.')
    parser.add_argument('--source', required=True, help='Path to the harvested bSDD TTL source file.')
    parser.add_argument('--cache', default=str(CACHE_JSON), help='Path to write the URI cache JSON.')
    args = parser.parse_args()
    cache_path = Path(args.cache)
    payload = build_cache(Path(args.source), cache_path)
    print(json.dumps({
        'source': payload['source'],
        'cache': str(cache_path),
        'uri_count': payload['uri_count'],
        'elapsed_seconds': payload['elapsed_seconds'],
        'cache_size_bytes': cache_path.stat().st_size,
    }, indent=2))
