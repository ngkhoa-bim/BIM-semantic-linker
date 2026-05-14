# =============================================================
# BIM Semantic Linker — Backend API v2.0
# Flask + IfcOpenShell + Neo4j Aura + Anthropic Claude
# =============================================================
# Endpoints:
#   GET  /health
#   POST /api/projects/<project_id>/ifc/ingest
#   GET  /api/projects/<project_id>/anchors
#   POST /api/projects/<project_id>/documents/analyze
#   POST /api/projects/<project_id>/links/confirm
#   GET  /api/projects/<project_id>/links
# =============================================================

import os
import json
import uuid
import hashlib
import re
from datetime import datetime, timezone

from flask import Flask, request, jsonify
from werkzeug.utils import secure_filename
from neo4j import GraphDatabase
import ifcopenshell
import ifcopenshell.util.element
import google.generativeai as genai

from dotenv import load_dotenv

load_dotenv()

# ──────────────────────────────────────────────────────────────
# App & Cấu hình
# ──────────────────────────────────────────────────────────────
app = Flask(__name__)

UPLOAD_FOLDER = os.environ.get("UPLOAD_FOLDER", "/tmp/ifc_uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

NEO4J_URI      = os.environ.get("NEO4J_URI",      "bolt://localhost:7687")
NEO4J_USER     = os.environ.get("NEO4J_USER",     "neo4j")
NEO4J_PASSWORD = os.environ.get("NEO4J_PASSWORD", "")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-1.5-flash")

_gemini_model = None

# ──────────────────────────────────────────────────────────────
# Lazy singletons (khởi tạo 1 lần, dùng lại)
# ──────────────────────────────────────────────────────────────
_neo4j_driver     = None
_anthropic_client = None


def get_driver():
    global _neo4j_driver
    if _neo4j_driver is None:
        _neo4j_driver = GraphDatabase.driver(
            NEO4J_URI, auth=(NEO4J_USER, NEO4J_PASSWORD)
        )
    return _neo4j_driver


def get_gemini_model():
    global _gemini_model
    if _gemini_model is None:
        if not GEMINI_API_KEY:
            raise RuntimeError("Missing GEMINI_API_KEY")
        genai.configure(api_key=GEMINI_API_KEY)
        _gemini_model = genai.GenerativeModel(GEMINI_MODEL)
    return _gemini_model


# ──────────────────────────────────────────────────────────────
# Neo4j — Tạo constraint & index khi khởi động
# ──────────────────────────────────────────────────────────────
def setup_constraints():
    """
    Tạo unique constraint và index tăng tốc truy vấn.
    Chạy 1 lần khi app khởi động, an toàn nếu đã tồn tại.
    """
    stmts = [
        "CREATE CONSTRAINT project_id_unique IF NOT EXISTS "
        "FOR (p:Project) REQUIRE p.id IS UNIQUE",

        "CREATE CONSTRAINT element_gid_unique IF NOT EXISTS "
        "FOR (e:Element) REQUIRE e.globalId IS UNIQUE",

        "CREATE CONSTRAINT document_id_unique IF NOT EXISTS "
        "FOR (d:Document) REQUIRE d.id IS UNIQUE",

        "CREATE INDEX element_name_idx IF NOT EXISTS "
        "FOR (e:Element) ON (e.name)",

        "CREATE INDEX element_project_idx IF NOT EXISTS "
        "FOR (e:Element) ON (e.project_id)",
    ]
    try:
        with get_driver().session() as s:
            for stmt in stmts:
                s.run(stmt)
        print("[OK] Neo4j constraints & indexes ready.")
    except Exception as e:
        # Không crash app nếu Neo4j chưa sẵn sàng ngay
        print(f"[WARN] setup_constraints: {e}")


# ──────────────────────────────────────────────────────────────
# IFC Helpers
# ──────────────────────────────────────────────────────────────
def safe_str(val) -> str:
    """Chuyển bất kỳ giá trị IFC nào thành string an toàn cho Neo4j."""
    if val is None:
        return ""
    if isinstance(val, (str, int, float, bool)):
        return str(val)
    # Xử lý IfcLabel, IfcIdentifier, enum, v.v.
    return str(val)


def flatten_psets(psets: dict) -> dict:
    """
    Làm phẳng nested psets thành dict 2 cấp:
    {pset_name: {prop_name: str_value}}
    Loại bỏ mọi object không serializable.
    """
    result = {}
    for pset_name, props in psets.items():
        if not isinstance(props, dict):
            continue
        flat = {}
        for k, v in props.items():
            flat[str(k)] = safe_str(v)
        if flat:
            result[str(pset_name)] = flat
    return result


def extract_mark_and_reference(psets: dict) -> tuple:
    """
    Trích xuất Mark và Reference từ các Pset phổ biến.
    Ưu tiên các pset chuẩn IFC trước, fallback sang tất cả psets.
    """
    mark = reference = ""

    # Các Pset chuẩn IFC chứa Mark/Reference
    standard_psets = [
        "Pset_BeamCommon", "Pset_ColumnCommon", "Pset_SlabCommon",
        "Pset_WallCommon", "Pset_DoorCommon", "Pset_WindowCommon",
        "Pset_MemberCommon", "Pset_PlateCommon", "Pset_StairCommon",
        "Pset_RampCommon", "Pset_CoveringCommon", "Pset_RoofCommon",
        "Pset_BuildingElementProxyCommon",
    ]

    for pset_name in standard_psets:
        if pset_name in psets and isinstance(psets[pset_name], dict):
            props = psets[pset_name]
            if not mark and props.get("Mark"):
                mark = safe_str(props["Mark"])
            if not reference and props.get("Reference"):
                reference = safe_str(props["Reference"])

    # Fallback: quét toàn bộ nếu chưa tìm được
    if not mark or not reference:
        for pset_name, props in psets.items():
            if not isinstance(props, dict):
                continue
            if not mark and props.get("Mark"):
                mark = safe_str(props["Mark"])
            if not reference and props.get("Reference"):
                reference = safe_str(props["Reference"])

    return mark, reference


def get_spatial_info(element) -> dict:
    """
    Lấy thông tin vị trí không gian: Storey, Building.
    Traverses IfcRelContainedInSpatialStructure.
    """
    info = {"storey": "", "building": ""}
    try:
        contained = getattr(element, "ContainedInStructure", None) or []
        for rel in contained:
            container = rel.RelatingStructure
            if container.is_a("IfcBuildingStorey"):
                info["storey"] = safe_str(container.Name)
            elif container.is_a("IfcBuilding"):
                info["building"] = safe_str(container.Name)
    except Exception:
        pass
    return info


def get_type_name(element) -> str:
    """
    Lấy tên kiểu cấu kiện từ IfcRelDefinesByType.
    Ví dụ: "300x600mm Concrete Beam"
    """
    try:
        defined_by = getattr(element, "IsDefinedBy", None) or []
        for rel in defined_by:
            if rel.is_a("IfcRelDefinesByType"):
                type_obj = rel.RelatingType
                if type_obj and type_obj.Name:
                    return safe_str(type_obj.Name)
    except Exception:
        pass
    return ""


# ──────────────────────────────────────────────────────────────
# Matching Helpers
# ──────────────────────────────────────────────────────────────
def id_based_match(doc_text: str, anchors: list) -> list:
    """
    ID-based matching thuần Python — KHÔNG cần LLM.
    Tìm kiếm chính xác (case-insensitive) các giá trị anchor trong văn bản.

    Thứ tự ưu tiên field (confidence cao → thấp):
      globalId(0.95) > mark(0.90) > reference(0.88) > tag(0.80) > name(0.75)
      > typeName(0.70) > objectType(0.65)

    Trả về: list of {anchor, matched_value, match_field, method, confidence}
    """
    text_lower = doc_text.lower()
    matched    = []
    seen_gids  = set()

    field_priority = [
        ("globalId",   0.95),
        ("mark",       0.90),
        ("reference",  0.88),
        ("tag",        0.80),
        ("name",       0.75),
        ("typeName",   0.70),
        ("objectType", 0.65),
    ]

    for anchor in anchors:
        gid = anchor.get("globalId", "")
        if not gid or gid in seen_gids:
            continue

        for field, confidence in field_priority:
            value = (anchor.get(field) or "").strip()
            # Bỏ qua giá trị quá ngắn (≤2 ký tự) — dễ false positive
            if len(value) < 3:
                continue
            if value.lower() in text_lower:
                matched.append({
                    "anchor":        anchor,
                    "matched_value": value,
                    "match_field":   field,
                    "method":        "ID_BASED",
                    "confidence":    confidence,
                })
                seen_gids.add(gid)
                break  # Chỉ lấy match ưu tiên cao nhất cho mỗi element

    return matched


def call_gemini_for_suggestions(
    doc_name: str,
    doc_text: str,
    anchors: list,
    id_matches: list,
) -> list:
    """
    Gemini Semantic Discovery:
    1. Xác nhận kết quả ID-based matching.
    2. Bổ sung liên kết ngữ nghĩa.
    3. Trả về JSON suggestions.
    """
    # Tóm tắt anchor list (giới hạn 150 để tránh context quá lớn)
    anchor_lines = []
    for a in anchors[:150]:
        parts = [f"{a.get('ifcType', '')}"]
        if a.get("name"):         parts.append(f"Name={a['name']}")
        if a.get("mark"):         parts.append(f"Mark={a['mark']}")
        if a.get("reference"):    parts.append(f"Ref={a['reference']}")
        if a.get("tag"):          parts.append(f"Tag={a['tag']}")
        if a.get("storey"):       parts.append(f"Storey={a['storey']}")
        if a.get("typeName"):     parts.append(f"Type={a['typeName']}")
        anchor_lines.append(f"  [{a['globalId']}] {', '.join(parts)}")

    # Tóm tắt ID-based matches để Claude xác nhận/bác bỏ
    id_match_lines = [
        f"  - [{m['anchor']['globalId']}] {m['anchor'].get('name','')} "
        f"(khớp qua {m['match_field']}='{m['matched_value']}')"
        for m in id_matches
    ] or ["  (Không có kết quả từ ID-based matching)"]

    # Giới hạn doc_text để tránh vượt context window
    doc_snippet = doc_text[:4000]
    if len(doc_text) > 4000:
        doc_snippet += "\n... [văn bản đã bị cắt bớt]"

    prompt = f"""Bạn là chuyên gia BIM (Building Information Modeling) với kinh nghiệm về IFC và Linked Building Data.

## TÊN TÀI LIỆU
{doc_name}

## NỘI DUNG TÀI LIỆU
{doc_snippet}

## DANH SÁCH CẤU KIỆN TRONG MÔ HÌNH BIM (anchor list)
{chr(10).join(anchor_lines)}

## CÁC LIÊN KẾT ĐÃ TÌM QUA ID-BASED MATCHING (cần xác nhận ngữ cảnh)
{chr(10).join(id_match_lines)}

## NHIỆM VỤ
1. XÁC NHẬN các liên kết ID-based trên nếu chúng thực sự có nghĩa trong ngữ cảnh tài liệu.
2. LOẠI BỎ nếu chỉ là trùng tên ngẫu nhiên, không có liên quan thực sự.
3. BỔ SUNG thêm các liên kết mới mà ID-based matching bỏ sót (tìm qua ngữ nghĩa, mô tả, vị trí, công năng).
4. Với mỗi đề xuất, TRÍCH DẪN CHÍNH XÁC đoạn văn bản làm bằng chứng (evidence).
5. Đánh giá confidence từ 0.0 đến 1.0.
6. Phân loại method: ID_BASED, SEMANTIC, hoặc HYBRID.

QUAN TRỌNG:
- Không được hallucinate GlobalId — chỉ dùng các GlobalId từ anchor list ở trên.
- Nếu không chắc, đặt confidence thấp (< 0.5) thay vì bỏ qua.
- Chỉ trả về JSON hợp lệ, không có text bên ngoài JSON.

OUTPUT FORMAT (JSON thuần, không markdown):
{{
  "suggestions": [
    {{
      "element_global_id": "string — GlobalId chính xác từ anchor list",
      "element_name": "string — tên cấu kiện",
      "matched_anchor": "string — giá trị đã khớp hoặc mô tả khớp ngữ nghĩa",
      "relationship_type": "REFERENCES",
      "confidence": 0.0,
      "method": "ID_BASED | SEMANTIC | HYBRID",
      "evidence": "string — trích dẫn chính xác từ tài liệu"
    }}
  ]
}}"""

    try:
        model = get_gemini_model()

        response = model.generate_content(
            prompt,
            generation_config={
                "temperature": 0.1,
                "top_p": 0.8,
                "top_k": 40,
                "max_output_tokens": 2048,
                "response_mime_type": "application/json",
            },
        )
        raw = response.text.strip()

        # Loại bỏ markdown code fence nếu Claude thêm vào
        raw = re.sub(r"^```(?:json)?\\s*", "", raw, flags=re.MULTILINE)
        raw = re.sub(r"\\s*```\\s*$", "", raw, flags=re.MULTILINE)

        parsed = json.loads(raw)
        return parsed.get("suggestions", [])

    except json.JSONDecodeError as e:
        print(f"[WARN] Gemini trả về JSON không hợp lệ: {e}")
        # Fallback: trả về ID-based matches dưới dạng suggestions
        return [
            {
                "element_global_id": m["anchor"]["globalId"],
                "element_name":      m["anchor"].get("name", ""),
                "matched_anchor":    m["matched_value"],
                "relationship_type": "REFERENCES",
                "confidence":        m["confidence"],
                "method":            "ID_BASED",
                "evidence":          f"ID-based: '{m['matched_value']}' found in document",
            }
            for m in id_matches
        ]

    except Exception as e:
        print(f"[ERROR] Gemini API call failed: {e}")
        return []


# ──────────────────────────────────────────────────────────────
# Endpoints
# ──────────────────────────────────────────────────────────────

@app.route("/health", methods=["GET"])
def health():
    """Health check — dùng để Ngrok/Frontend kiểm tra backend đang chạy."""
    return jsonify({
        "status":    "ok",
        "service":   "BIM Semantic Linker Backend",
        "version":   "2.0.0",
        "timestamp": datetime.now(timezone.utc).isoformat(),
    })


@app.route("/api/projects/<project_id>/ifc/ingest", methods=["POST"])
def ingest_ifc(project_id: str):
    """
    Nhận file IFC qua multipart/form-data.
    Trích xuất cấu kiện và nạp vào Neo4j theo cấu trúc:
      (Project)-[:HAS_ELEMENT]->(Element)
                                    └-[:HAS_PSET]->(PropertySet)
                                                       └-[:HAS_PROPERTY]->(Property)
    """
    if "file" not in request.files:
        return jsonify({"error": "Thiếu trường 'file' trong form-data"}), 400

    f = request.files["file"]
    if not f.filename:
        return jsonify({"error": "Tên file rỗng"}), 400

    filename = secure_filename(f.filename)
    tmp_path = os.path.join(UPLOAD_FOLDER, f"{uuid.uuid4().hex}_{filename}")

    try:
        f.save(tmp_path)
        ifc = ifcopenshell.open(tmp_path)

        # Lấy tất cả IfcBuildingElement (bao gồm mọi subtype)
        elements = ifc.by_type("IfcBuildingElement")

        ingested_count = 0
        errors         = []

        with get_driver().session() as session:
            # Tạo hoặc cập nhật Project node
            session.run(
                """
                MERGE (p:Project {id: $pid})
                SET p.name       = $pid,
                    p.updated_at = $ts
                """,
                pid=project_id,
                ts=datetime.now(timezone.utc).isoformat(),
            )

            for elem in elements:
                try:
                    gid           = elem.GlobalId
                    name          = safe_str(elem.Name)
                    ifc_type      = elem.is_a()
                    object_type   = safe_str(getattr(elem, "ObjectType", ""))
                    tag           = safe_str(getattr(elem, "Tag", ""))
                    predefined_type = safe_str(getattr(elem, "PredefinedType", ""))

                    # Trích xuất Psets, Mark, Reference, Spatial, TypeName
                    psets_raw         = ifcopenshell.util.element.get_psets(elem)
                    mark, reference   = extract_mark_and_reference(psets_raw)
                    spatial           = get_spatial_info(elem)
                    type_name         = get_type_name(elem)
                    psets_flat        = flatten_psets(psets_raw)

                    # ── Tạo Element node ──────────────────────────────────
                    session.run(
                        """
                        MERGE (e:Element {globalId: $gid})
                        SET e.name            = $name,
                            e.ifcType         = $ifc_type,
                            e.objectType      = $object_type,
                            e.tag             = $tag,
                            e.predefinedType  = $predefined_type,
                            e.mark            = $mark,
                            e.reference       = $reference,
                            e.storey          = $storey,
                            e.typeName        = $type_name,
                            e.project_id      = $pid
                        WITH e
                        MATCH (p:Project {id: $pid})
                        MERGE (p)-[:HAS_ELEMENT]->(e)
                        """,
                        gid=gid, name=name, ifc_type=ifc_type,
                        object_type=object_type, tag=tag,
                        predefined_type=predefined_type, mark=mark,
                        reference=reference, storey=spatial["storey"],
                        type_name=type_name, pid=project_id,
                    )

                    # ── Tạo PropertySet + Property nodes ──────────────────
                    for pset_name, props in psets_flat.items():
                        # ID duy nhất cho PropertySet
                        pset_id = hashlib.md5(
                            f"{gid}::{pset_name}".encode()
                        ).hexdigest()

                        session.run(
                            """
                            MATCH (e:Element {globalId: $gid})
                            MERGE (ps:PropertySet {id: $psid})
                            SET ps.name = $pset_name
                            MERGE (e)-[:HAS_PSET]->(ps)
                            """,
                            gid=gid, psid=pset_id, pset_name=pset_name,
                        )

                        for prop_name, prop_val in props.items():
                            prop_id = hashlib.md5(
                                f"{pset_id}::{prop_name}".encode()
                            ).hexdigest()

                            session.run(
                                """
                                MATCH (ps:PropertySet {id: $psid})
                                MERGE (pr:Property {id: $prop_id})
                                SET pr.name  = $prop_name,
                                    pr.value = $prop_val
                                MERGE (ps)-[:HAS_PROPERTY]->(pr)
                                """,
                                psid=pset_id, prop_id=prop_id,
                                prop_name=prop_name, prop_val=prop_val,
                            )

                    ingested_count += 1

                except Exception as elem_err:
                    errors.append({
                        "globalId": getattr(elem, "GlobalId", "unknown"),
                        "error":    str(elem_err),
                    })

        return jsonify({
            "status":          "ok",
            "project_id":      project_id,
            "ingested_count":  ingested_count,
            "error_count":     len(errors),
            "errors":          errors[:10],  # Chỉ trả 10 lỗi đầu
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        # Luôn xoá file tạm dù có lỗi hay không
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.route("/api/projects/<project_id>/anchors", methods=["GET"])
def get_anchors(project_id: str):
    """
    Trả về danh sách anchor (định danh cấu kiện) từ Neo4j.
    Dùng để Frontend preview hoặc Backend dùng nội bộ trong analyze.
    """
    try:
        with get_driver().session() as session:
            result = session.run(
                """
                MATCH (p:Project {id: $pid})-[:HAS_ELEMENT]->(e:Element)
                RETURN e.globalId      AS globalId,
                       e.name          AS name,
                       e.ifcType       AS ifcType,
                       e.objectType    AS objectType,
                       e.tag           AS tag,
                       e.mark          AS mark,
                       e.reference     AS reference,
                       e.typeName      AS typeName,
                       e.storey        AS storey,
                       e.predefinedType AS predefinedType
                ORDER BY e.name
                """,
                pid=project_id,
            )
            anchors = [
                {
                    "globalId":       rec["globalId"]      or "",
                    "name":           rec["name"]          or "",
                    "ifcType":        rec["ifcType"]       or "",
                    "objectType":     rec["objectType"]    or "",
                    "tag":            rec["tag"]           or "",
                    "mark":           rec["mark"]          or "",
                    "reference":      rec["reference"]     or "",
                    "typeName":       rec["typeName"]      or "",
                    "storey":         rec["storey"]        or "",
                    "predefinedType": rec["predefinedType"] or "",
                }
                for rec in result
            ]

        return jsonify({
            "project_id": project_id,
            "count":      len(anchors),
            "anchors":    anchors,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/projects/<project_id>/documents/analyze", methods=["POST"])
def analyze_document(project_id: str):
    """
    Pipeline phân tích tài liệu dự án:
    1. Nhận document_name + document_text
    2. Lấy anchor list từ Neo4j
    3. Chạy ID-based matching (Python, không LLM)
    4. Gọi Claude cho Semantic Discovery + Evidence
    5. Trả về danh sách suggestions có schema chuẩn

    Body JSON:
    {
      "document_name": "SNT-BOQ-STR-ALL-v1.xlsx",
      "document_text": "..."
    }
    """
    data     = request.get_json(force=True)
    doc_name = (data.get("document_name") or "unknown.pdf").strip()
    doc_text = (data.get("document_text") or "").strip()

    if not doc_text:
        return jsonify({"error": "document_text không được rỗng"}), 400

    # Tạo document_id ổn định: hash(project_id + doc_name)
    doc_id = hashlib.sha256(
        f"{project_id}::{doc_name}".encode()
    ).hexdigest()[:16]

    # ── Bước 1: Lấy anchors từ Neo4j ─────────────────────────
    try:
        with get_driver().session() as session:
            result = session.run(
                """
                MATCH (p:Project {id: $pid})-[:HAS_ELEMENT]->(e:Element)
                RETURN e.globalId AS globalId, e.name AS name,
                       e.ifcType AS ifcType, e.objectType AS objectType,
                       e.tag AS tag, e.mark AS mark,
                       e.reference AS reference, e.typeName AS typeName,
                       e.storey AS storey, e.predefinedType AS predefinedType
                """,
                pid=project_id,
            )
            anchors = [dict(rec) for rec in result]
    except Exception as e:
        return jsonify({"error": f"Neo4j query failed: {e}"}), 500

    if not anchors:
        return jsonify({
            "document":    {"document_id": doc_id, "name": doc_name},
            "suggestions": [],
            "message":     (
                f"Project '{project_id}' chưa có element nào. "
                "Hãy ingest file IFC trước."
            ),
        })

    # ── Bước 2: ID-based matching (Python, không tốn token LLM) ──
    id_matches = id_based_match(doc_text, anchors)

    # ── Bước 3: Claude cho Semantic Discovery + Evidence ─────────
    gemini_suggestions = call_gemini_for_suggestions(
    doc_name, doc_text, anchors, id_matches
    )

    # ── Bước 4: Chuẩn hoá + gắn suggestion_id ────────────────────
    suggestions = []
    for s in gemini_suggestions:
        # Validate: element_global_id phải tồn tại trong anchor list
        valid_gids = {a["globalId"] for a in anchors}
        gid = s.get("element_global_id", "")
        if gid and gid not in valid_gids:
            # Claude hallucinated GlobalId không tồn tại — bỏ qua
            print(f"[WARN] Claude returned invalid GlobalId: {gid} — skipping")
            continue

        suggestions.append({
            "suggestion_id":    str(uuid.uuid4()),
            "element_global_id": gid,
            "element_name":      s.get("element_name", ""),
            "matched_anchor":    s.get("matched_anchor", ""),
            "relationship_type": s.get("relationship_type", "REFERENCES"),
            "confidence":        round(float(s.get("confidence", 0.5)), 3),
            "method":            s.get("method", "SEMANTIC"),
            "evidence":          s.get("evidence", ""),
            "status":            "PENDING",
        })

    return jsonify({
        "document": {
            "document_id": doc_id,
            "name":        doc_name,
        },
        "suggestions": suggestions,
    })


@app.route("/api/projects/<project_id>/links/confirm", methods=["POST"])
def confirm_link(project_id: str):
    """
    Ghi liên kết đã được con người xác nhận vào Neo4j.
    Tạo:
      (d:Document)-[:REFERENCES {metadata}]->(e:Element)

    Body JSON:
    {
      "document_id":       "abc123",
      "document_name":     "SNT-BOQ-STR-ALL-v1.xlsx",
      "element_global_id": "0ABC...",
      "suggestion_id":     "uuid",
      "confidence":        0.87,
      "method":            "HYBRID",
      "evidence":          "...",
      "confirmed_by":      "human_via_ui"   (optional)
    }
    """
    data = request.get_json(force=True)

    required_fields = ["document_id", "document_name", "element_global_id"]
    missing = [f for f in required_fields if not data.get(f)]
    if missing:
        return jsonify({"error": f"Thiếu trường: {', '.join(missing)}"}), 400

    doc_id       = data["document_id"]
    doc_name     = data["document_name"]
    gid          = data["element_global_id"]
    confidence   = round(float(data.get("confidence", 0.5)), 3)
    method       = data.get("method", "UNKNOWN")
    evidence     = data.get("evidence", "")
    confirmed_by = data.get("confirmed_by", "human")
    created_at   = datetime.now(timezone.utc).isoformat()

    try:
        with get_driver().session() as session:
            # Kiểm tra Element tồn tại trong project
            rec = session.run(
                """
                MATCH (e:Element {globalId: $gid, project_id: $pid})
                RETURN e.name AS name
                """,
                gid=gid, pid=project_id,
            ).single()

            if not rec:
                return jsonify({
                    "error": (
                        f"Không tìm thấy Element với globalId='{gid}' "
                        f"trong project '{project_id}'"
                    )
                }), 404

            element_name = rec["name"] or gid

            # Ghi Document node và REFERENCES relationship
            # MERGE trên relationship key (doc_name + gid) để tránh duplicate
            session.run(
                """
                MERGE (d:Document {id: $doc_id})
                SET d.name       = $doc_name,
                    d.project_id = $pid

                WITH d
                MATCH (e:Element {globalId: $gid})
                MERGE (d)-[r:REFERENCES {
                    source_document_name: $doc_name,
                    element_global_id:    $gid
                }]->(e)
                SET r.confidence  = $confidence,
                    r.method      = $method,
                    r.evidence    = $evidence,
                    r.created_at  = $created_at,
                    r.created_by  = $confirmed_by,
                    r.status      = "CONFIRMED"
                """,
                doc_id=doc_id, doc_name=doc_name, pid=project_id,
                gid=gid, confidence=confidence, method=method,
                evidence=evidence, created_at=created_at,
                confirmed_by=confirmed_by,
            )

        return jsonify({
            "status":            "ok",
            "message":           f"Đã tạo liên kết: '{doc_name}' → '{element_name}'",
            "document_id":       doc_id,
            "element_global_id": gid,
            "element_name":      element_name,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/projects/<project_id>/links", methods=["GET"])
def get_links(project_id: str):
    """
    Trả về tất cả liên kết đã xác nhận trong một project.
    Dùng cho tab 3 của Frontend.
    """
    try:
        with get_driver().session() as session:
            result = session.run(
                """
                MATCH (d:Document {project_id: $pid})-[r:REFERENCES]->(e:Element {project_id: $pid})
                RETURN d.name          AS document_name,
                       d.id            AS document_id,
                       e.name          AS element_name,
                       e.globalId      AS element_global_id,
                       e.ifcType       AS element_type,
                       r.confidence    AS confidence,
                       r.method        AS method,
                       r.evidence      AS evidence,
                       r.created_at    AS created_at,
                       r.created_by    AS created_by,
                       r.status        AS status
                ORDER BY r.created_at DESC
                """,
                pid=project_id,
            )
            links = [dict(rec) for rec in result]

        return jsonify({
            "project_id": project_id,
            "count":      len(links),
            "links":      links,
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ──────────────────────────────────────────────────────────────
# Khởi động
# ──────────────────────────────────────────────────────────────
with app.app_context():
    setup_constraints()

if __name__ == "__main__":
    # Port 7860 để tương thích Ngrok
    app.run(host="0.0.0.0", port=7860, debug=False)
