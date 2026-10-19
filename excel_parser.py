"""
Excel Data Model Parser & Field Mapper
Parses Abacus360 and AR Client Excel files, extracts table/field metadata,
and computes field mappings based on description similarity.
"""

import os
import re
import openpyxl
import pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

DATA_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "data")

ABACUS_FILE = os.path.join(DATA_DIR, "Abacus360_client.xlsx")
AR_FILE = os.path.join(DATA_DIR, "AR Client Fildes.xlsx")


# ─── Abacus360 Parsing ───────────────────────────────────────────────

def parse_abacus():
    """
    Returns dict with:
      tables: list of { name, component, caption, description, status, version }
      fields: list of { field, domain, component, caption, description, status, version }
      entity_fields: list of { entity, component, field, pKey, stored }
      relationships: list of { from_table, to_table, field, pKey }
    """

    wb = openpyxl.load_workbook(ABACUS_FILE, read_only=True)

    # --- Entities sheet ---
    ws = wb["Entities"]
    tables = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        tables.append({
            "name": str(row[0]).strip(),
            "component": row[1],
            "caption": row[3] or row[2] or "",
            "description": row[5] or row[4] or "",
            "status": row[6] or "",
            "version": row[7] or "",
        })

    # --- Fields sheet ---
    ws = wb["Fields"]
    fields = []
    field_lookup = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        f = {
            "field": str(row[0]).strip(),
            "domain": row[1] or "",
            "component": row[2] or "",
            "caption": row[4] or row[3] or "",
            "description": row[6] or row[5] or "",
            "status": row[8] or "",
            "version": row[9] or "",
        }
        fields.append(f)
        field_lookup[str(row[0]).strip()] = f

    # --- EntityFields sheet ---
    ws = wb["EntityFields"]
    entity_fields = []
    # Build relationships: entity -> fields
    entity_field_map = {}  # entity_name -> list of field dicts
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        ef = {
            "entity": str(row[0]).strip(),
            "component": row[1] or "",
            "field": str(row[2]).strip(),
            "pKey": str(row[3]).lower() == "true" if row[3] else False,
            "stored": str(row[4]).lower() == "true" if row[4] else False,
            "status": row[5] or "",
            "version": row[6] or "",
        }
        entity_fields.append(ef)
        entity_field_map.setdefault(str(row[0]).strip(), []).append(ef)

    wb.close()

    # Build enriched table data (tables with their fields)
    tables_enriched = []
    for t in tables:
        t_fields = entity_field_map.get(t["name"], [])
        enriched_fields = []
        for ef in t_fields:
            fl = field_lookup.get(ef["field"], {})
            enriched_fields.append({
                "field": ef["field"],
                "pKey": ef["pKey"],
                "stored": ef["stored"],
                "caption": fl.get("caption", ""),
                "description": fl.get("description", ""),
                "domain": fl.get("domain", ""),
                "type": ef.get("status", ""),
            })
        tables_enriched.append({
            **t,
            "fields": enriched_fields,
            "field_count": len(enriched_fields),
        })

    return {
        "tables": tables_enriched,
        "fields": fields,
        "entity_fields": entity_fields,
    }


# ─── AR Client Parsing ───────────────────────────────────────────────

def parse_ar_client():
    """
    Returns dict with:
      tables: list of { name, alias, access_group, fields: [...] }
    """
    wb = openpyxl.load_workbook(AR_FILE, read_only=True)
    ws = wb.active

    table_map = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[0]:
            continue
        tname = str(row[0]).strip()
        if tname not in table_map:
            table_map[tname] = {
                "name": tname,
                "alias": row[1] or "",
                "access_group": row[6] or "",
                "fields": [],
            }
        table_map[tname]["fields"].append({
            "field": row[2] or "",
            "caption": row[3] or "",
            "type": row[4] or "",
            "size": row[5] or "",
            "description": row[7] or "",
            "variables": row[8] or "",
        })

    wb.close()

    tables = list(table_map.values())
    for t in tables:
        t["field_count"] = len(t["fields"])
    return {"tables": tables}


# ─── Field Mapping (TF-IDF + Cosine Similarity) ──────────────────────

def _clean_text(text):
    """Lowercase, remove special chars, collapse whitespace."""
    if not text:
        return ""
    text = str(text).lower()
    text = re.sub(r"[^a-z0-9\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def compute_field_mapping(top_n=5, threshold=0.1):
    """
    For every AR Client field, find the top_n most similar Abacus360 fields
    based on TF-IDF cosine similarity of names, captions, descriptions, and type metadata.
    Returns list of { ar_table, ar_field, ar_caption, ar_description,
                      ar_type, ar_size,
                      abacus_entity, abacus_field, abacus_caption, abacus_description,
                      abacus_type, score }
    """
    abacus = parse_abacus()
    ar = parse_ar_client()

    abacus_entities_by_field = {}

    # Build flat lists of (meta, text) for both sides
    abacus_items = []
    abacus_texts = []
    for table in abacus["tables"]:
        for f in table["fields"]:
            abacus_type = f.get("type", "")
            abacus_entities_by_field.setdefault(f["field"], set()).add(table["name"])
            combined = " ".join(
                part for part in [
                    _clean_text(f["field"]),
                    _clean_text(f["caption"]),
                    _clean_text(f["description"]),
                    _clean_text(abacus_type),
                ]
                if part
            )
            if not combined.strip():
                continue
            abacus_items.append({
                "entity": table["name"],
                "field": f["field"],
                "caption": f["caption"],
                "description": f["description"],
                "type": abacus_type,
            })
            abacus_texts.append(combined)

    ar_items = []
    ar_texts = []
    for table in ar["tables"]:
        for f in table["fields"]:
            combined = " ".join(
                part for part in [
                    _clean_text(f["field"]),
                    _clean_text(f["caption"]),
                    _clean_text(f["description"]),
                    _clean_text(f.get("variables", "")),
                    _clean_text(f.get("type", "")),
                    _clean_text(f.get("size", "")),
                ]
                if part
            )
            if not combined.strip():
                continue
            ar_items.append({
                "table": table["name"],
                "field": f["field"],
                "caption": f["caption"],
                "description": f["description"],
                "type": f.get("type", ""),
                "size": f.get("size", ""),
            })
            ar_texts.append(combined)

    if not abacus_texts or not ar_texts:
        return []

    # TF-IDF vectorization
    vectorizer = TfidfVectorizer(max_features=10000, stop_words="english")
    all_texts = abacus_texts + ar_texts
    tfidf_matrix = vectorizer.fit_transform(all_texts)

    abacus_tfidf = tfidf_matrix[: len(abacus_texts)]
    ar_tfidf = tfidf_matrix[len(abacus_texts):]

    # Cosine similarity: ar_fields x abacus_fields
    sim_matrix = cosine_similarity(ar_tfidf, abacus_tfidf)

    mappings = []
    for i, ar_item in enumerate(ar_items):
        scores = sim_matrix[i]
        top_indices = scores.argsort()[::-1] if top_n is None else scores.argsort()[::-1][:top_n]
        for j in top_indices:
            score = float(scores[j])
            if score < threshold:
                continue
            mappings.append({
                "ar_table": ar_item["table"],
                "ar_field": ar_item["field"],
                "ar_caption": ar_item["caption"],
                "ar_description": ar_item["description"],
                "ar_type": ar_item["type"],
                "ar_size": ar_item["size"],
                "abacus_entity": abacus_items[j]["entity"],
                "abacus_field": abacus_items[j]["field"],
                "abacus_caption": abacus_items[j]["caption"],
                "abacus_description": abacus_items[j]["description"],
                "abacus_type": abacus_items[j]["type"],
                "abacus_entities": sorted(abacus_entities_by_field.get(abacus_items[j]["field"], {abacus_items[j]["entity"]})),
                "score": round(score, 4),
            })

    # Sort by score descending
    mappings.sort(key=lambda x: x["score"], reverse=True)
    return mappings
