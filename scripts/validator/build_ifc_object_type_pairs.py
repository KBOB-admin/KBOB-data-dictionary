from __future__ import annotations

import argparse
import json
from pathlib import Path

from rdflib import Graph, OWL, RDF, RDFS, URIRef


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TARGET = REPO_ROOT / 'resources' / 'ifc' / 'ifc4x3-add2-object-type-pairs.json'
IFC_NAMESPACE = 'https://standards.buildingsmart.org/IFC/DEV/IFC4x3/ADD2/OWL#'


def _is_subclass_of(graph: Graph, candidate: URIRef, ancestor: URIRef) -> bool:
    visited: set[URIRef] = set()
    pending = [candidate]
    while pending:
        current = pending.pop()
        if current == ancestor:
            return True
        if current in visited:
            continue
        visited.add(current)
        pending.extend(
            parent
            for parent in graph.objects(current, RDFS.subClassOf)
            if isinstance(parent, URIRef)
        )
    return False


def build_pairs(source: Path, target: Path = DEFAULT_TARGET) -> dict:
    if not source.exists():
        raise FileNotFoundError(f'IFC4X3_ADD2 ontology not found: {source}')

    graph = Graph()
    graph.parse(source, format='turtle')
    classes = {
        entity
        for entity in graph.subjects(RDF.type, OWL.Class)
        if isinstance(entity, URIRef) and str(entity).startswith(IFC_NAMESPACE)
    }
    ifc_object = URIRef(f'{IFC_NAMESPACE}IfcObject')
    ifc_type_object = URIRef(f'{IFC_NAMESPACE}IfcTypeObject')
    pairs: dict[str, str] = {}

    for object_entity in classes:
        object_name = str(object_entity).removeprefix(IFC_NAMESPACE)
        type_name = f'{object_name}Type'
        type_entity = URIRef(f'{IFC_NAMESPACE}{type_name}')
        if type_entity not in classes:
            continue
        if not _is_subclass_of(graph, object_entity, ifc_object):
            continue
        if not _is_subclass_of(graph, type_entity, ifc_type_object):
            continue
        pairs[object_name] = type_name

    payload = {
        'ifc_release': 'IFC4X3_ADD2',
        'source_ontology': 'https://standards.buildingsmart.org/IFC/DEV/IFC4x3/ADD2/OWL',
        'pair_count': len(pairs),
        'object_type_pairs': dict(sorted(pairs.items())),
    }
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    return payload


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Build IFC object/TypeObject pairs from the IFC4X3_ADD2 ontology.')
    parser.add_argument('--source', required=True, help='Path to the IFC4X3_ADD2 Turtle ontology.')
    parser.add_argument('--target', default=str(DEFAULT_TARGET), help='Output JSON path.')
    args = parser.parse_args()
    result = build_pairs(Path(args.source), Path(args.target))
    print(json.dumps({'target': args.target, 'pair_count': result['pair_count']}, indent=2))
