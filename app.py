"""
Flask API for Excel Data Model Viewer & Field Mapper
"""

import io
from flask import Flask, jsonify, request, send_file
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from excel_parser import parse_abacus, parse_ar_client, compute_field_mapping

app = Flask(__name__)
CORS(app)

# Cache parsed data in memory
_cache = {}


def _get_abacus():
    if "abacus" not in _cache:
        _cache["abacus"] = parse_abacus()
    return _cache["abacus"]


def _get_ar():
    if "ar" not in _cache:
        _cache["ar"] = parse_ar_client()
    return _cache["ar"]


def _get_mappings():
    if "mappings" not in _cache:
        # Keep a bounded candidate set per AR field; the UI groups duplicates and expands same-field entities.
        _cache["mappings"] = compute_field_mapping(top_n=3, threshold=0.0)
    return _cache["mappings"]


# ─── Abacus360 endpoints ─────────────────────────────────────────────

@app.route("/api/abacus/tables", methods=["GET"])
def abacus_tables():
    """Return all Abacus360 tables (entities) with field counts."""
    data = _get_abacus()
    # Return tables without full field lists for overview
    tables = []
    for t in data["tables"]:
        tables.append({
            "name": t["name"],
            "component": t["component"],
            "caption": t["caption"],
            "description": t["description"],
            "status": t["status"],
            "version": t["version"],
            "field_count": t["field_count"],
        })
    return jsonify(tables)


@app.route("/api/abacus/tables/<table_name>", methods=["GET"])
def abacus_table_detail(table_name):
    """Return a single Abacus360 table with all its fields."""
    data = _get_abacus()
    for t in data["tables"]:
        if t["name"] == table_name:
            return jsonify(t)
    return jsonify({"error": "Table not found"}), 404


@app.route("/api/abacus/er-diagram", methods=["GET"])
def abacus_er_diagram():
    """
    Return ER diagram data: tables with their fields and relationships
    derived from shared fields between entities.
    """
    data = _get_abacus()
    tables = data["tables"]

    # Build nodes
    nodes = []
    for t in tables:
        nodes.append({
            "id": t["name"],
            "caption": t["caption"],
            "component": t["component"],
            "fields": [
                {
                    "name": f["field"],
                    "caption": f["caption"],
                    "pKey": f["pKey"],
                    "domain": f["domain"],
                }
                for f in t["fields"]
            ],
        })

    # Build edges from shared fields between entities
    # A shared field = same field name used in 2+ entities (foreign key pattern)
    field_to_entities = {}
    for t in tables:
        for f in t["fields"]:
            field_to_entities.setdefault(f["field"], set()).add(t["name"])

    edges = []
    for field, entities in field_to_entities.items():
        elist = sorted(entities)
        if len(elist) < 2:
            continue
        # Find which entity has it as PK (parent) vs non-PK (child)
        pk_entities = set()
        for t in tables:
            for f in t["fields"]:
                if f["field"] == field and f["pKey"]:
                    pk_entities.add(t["name"])

        for i in range(len(elist)):
            for j in range(i + 1, len(elist)):
                # Direction: from parent (PK holder) to child
                if elist[i] in pk_entities and elist[j] not in pk_entities:
                    edges.append({"from": elist[i], "to": elist[j], "shared_field": field, "type": "fk"})
                elif elist[j] in pk_entities and elist[i] not in pk_entities:
                    edges.append({"from": elist[j], "to": elist[i], "shared_field": field, "type": "fk"})
                else:
                    edges.append({"from": elist[i], "to": elist[j], "shared_field": field, "type": "shared"})

    return jsonify({"nodes": nodes, "edges": edges})


# ─── AR Client endpoints ─────────────────────────────────────────────

@app.route("/api/ar/tables", methods=["GET"])
def ar_tables():
    """Return all AR Client tables with field counts."""
    data = _get_ar()
    tables = []
    for t in data["tables"]:
        tables.append({
            "name": t["name"],
            "alias": t["alias"],
            "access_group": t["access_group"],
            "field_count": t["field_count"],
        })
    return jsonify(tables)


@app.route("/api/ar/tables/<table_name>", methods=["GET"])
def ar_table_detail(table_name):
    """Return a single AR Client table with all its fields."""
    data = _get_ar()
    for t in data["tables"]:
        if t["name"] == table_name:
            return jsonify(t)
    return jsonify({"error": "Table not found"}), 404


@app.route("/api/ar/er-diagram", methods=["GET"])
def ar_er_diagram():
    """Return AR Client ER diagram data."""
    data = _get_ar()
    nodes = []
    for t in data["tables"]:
        nodes.append({
            "id": t["name"],
            "caption": t["alias"],
            "fields": [
                {
                    "name": f["field"],
                    "caption": f["caption"],
                    "type": f["type"],
                    "size": f["size"],
                }
                for f in t["fields"]
            ],
        })
    # Build edges from shared field names between AR tables
    field_to_tables = {}
    for t in data["tables"]:
        for f in t["fields"]:
            field_to_tables.setdefault(f["field"], set()).add(t["name"])

    edges = []
    for field, tnames in field_to_tables.items():
        tlist = sorted(tnames)
        if len(tlist) < 2:
            continue
        for i in range(len(tlist)):
            for j in range(i + 1, len(tlist)):
                edges.append({"from": tlist[i], "to": tlist[j], "shared_field": field, "type": "shared"})

    return jsonify({"nodes": nodes, "edges": edges})


# ─── Mapping endpoints ───────────────────────────────────────────────

@app.route("/api/mappings", methods=["GET"])
def field_mappings():
    """
    Return field mappings between AR Client and Abacus360
    based on description similarity.
    Query params:
      - ar_table: filter by AR table name
      - abacus_entity: filter by Abacus entity name
      - min_score: minimum similarity score (0-1), default 0.1
      - limit: max results, default 500
    """
    mappings = _get_mappings()

    ar_table = request.args.get("ar_table")
    abacus_entity = request.args.get("abacus_entity")
    min_score = float(request.args.get("min_score", 0.1))
    limit = request.args.get("limit")
    limit = int(limit) if limit not in (None, "") else None

    filtered = mappings
    if ar_table:
        filtered = [m for m in filtered if m["ar_table"] == ar_table]
    if abacus_entity:
        filtered = [m for m in filtered if m["abacus_entity"] == abacus_entity]
    filtered = [m for m in filtered if m["score"] >= min_score]

    return jsonify(filtered if limit is None else filtered[:limit])


@app.route("/api/mappings/summary", methods=["GET"])
def mapping_summary():
    """Return summary: how many mappings per AR table / Abacus entity pair."""
    mappings = _get_mappings()
    min_score = float(request.args.get("min_score", 0.0))

    pair_counts = {}
    for m in mappings:
        if m["score"] < min_score:
            continue
        key = (m["ar_table"], m["abacus_entity"])
        pair_counts[key] = pair_counts.get(key, 0) + 1

    summary = [
        {"ar_table": k[0], "abacus_entity": k[1], "mapping_count": v}
        for k, v in sorted(pair_counts.items(), key=lambda x: x[1], reverse=True)
    ]
    return jsonify(summary)


# ─── Export endpoints ─────────────────────────────────────────────────

@app.route("/api/mappings/export", methods=["GET"])
def export_mappings_excel():
    """Export field mappings as a styled Excel file."""
    mappings = _get_mappings()

    ar_table = request.args.get("ar_table")
    abacus_entity = request.args.get("abacus_entity")
    min_score = float(request.args.get("min_score", 0.1))

    filtered = mappings
    if ar_table:
        filtered = [m for m in filtered if m["ar_table"] == ar_table]
    if abacus_entity:
        filtered = [m for m in filtered if m["abacus_entity"] == abacus_entity]
    filtered = [m for m in filtered if m["score"] >= min_score]

    # Create workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Field Mappings"

    # Styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin")
    )
    wrap_align = Alignment(vertical="top", wrap_text=True)

    headers = [
        "AR Table", "AR Field", "AR Type", "AR Size", "AR Caption", "AR Description",
        "Similarity Score", "Score %",
        "Abacus Entity", "Abacus Field", "Abacus Type", "Abacus Caption", "Abacus Description"
    ]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # Data rows
    for i, m in enumerate(filtered, 2):
        row_data = [
            m["ar_table"], m["ar_field"], m.get("ar_type", ""), m.get("ar_size", ""), m["ar_caption"], m["ar_description"],
            m["score"], f"{m['score'] * 100:.1f}%",
            m["abacus_entity"], m["abacus_field"], m.get("abacus_type", ""), m["abacus_caption"], m["abacus_description"]
        ]
        for col, val in enumerate(row_data, 1):
            cell = ws.cell(row=i, column=col, value=val)
            cell.border = thin_border
            cell.alignment = wrap_align

        # Color-code score
        score_cell = ws.cell(row=i, column=7)
        if m["score"] >= 0.5:
            score_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif m["score"] >= 0.3:
            score_cell.fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
        else:
            score_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Column widths
    widths = [16, 22, 12, 12, 25, 45, 14, 10, 20, 18, 14, 25, 45]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    # Freeze header row
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Write to buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)

    return send_file(
        buffer,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="field_mappings.xlsx"
    )


# ─── Cache management ────────────────────────────────────────────────

@app.route("/api/refresh", methods=["POST"])
def refresh_cache():
    """Clear cache and re-parse files."""
    _cache.clear()
    return jsonify({"status": "ok"})


if __name__ == "__main__":
    app.run(debug=True, port=5000)
