import csv
import io
import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple, Any

from flask import Flask, Response, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename
from openpyxl import load_workbook

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = Path(os.environ.get("MATERIAL_DATA_DIR", BASE_DIR / "data"))
UPLOAD_DIR = DATA_DIR / "uploads"
DB_PATH = DATA_DIR / "materials.db"
ALLOWED_EXTENSIONS = {"xlsx", "xlsm"}

app = Flask(__name__)
app.secret_key = os.environ.get("MATERIAL_APP_SECRET", "change-this-secret")
DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

STATUS_OPTIONS = [
    "Chưa đặt hàng",
    "Đã gửi yêu cầu báo giá",
    "Đã đặt hàng",
    "Đang sản xuất/chuẩn bị",
    "Đang vận chuyển",
    "Đã về kho",
    "Đã bàn giao công trình",
    "Thiếu / Chờ bổ sung",
    "Hủy / Không mua",
]

HEADER_ALIASES = {
    "item_code": ["ma vat tu", "ma hang", "code", "item code", "material code", "ma", "tag"],
    "item_name": ["ten vat tu", "ten hang", "noi dung", "description", "item", "material", "ten thiet bi", "hang hoa"],
    "specification": ["quy cach", "thong so", "spec", "specification", "model", "kich thuoc", "vat lieu"],
    "unit": ["don vi", "dvt", "unit", "uom"],
    "quantity": ["so luong", "sl", "qty", "quantity", "khoi luong"],
    "supplier": ["nha cung cap", "supplier", "vendor", "hang sx", "manufacturer"],
    "po_no": ["po", "po no", "so po", "purchase order", "don hang"],
    "request_date": ["ngay yeu cau", "request date", "ngay dat", "order date", "ngay mua"],
    "required_date": ["ngay can", "required date", "delivery date", "ngay giao", "eta", "deadline"],
    "note": ["ghi chu", "note", "remark", "remarks"],
}


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    replacements = str.maketrans(
        "àáạảãâầấậẩẫăằắặẳẵèéẹẻẽêềếệểễìíịỉĩòóọỏõôồốộổỗơờớợởỡùúụủũưừứựửữỳýỵỷỹđ",
        "aaaaaaaaaaaaaaaaaeeeeeeeeeeeiiiiiooooooooooooooooouuuuuuuuuuuyyyyyd",
    )
    return " ".join(text.translate(replacements).replace("_", " ").replace("/", " ").split())


def db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with db() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                code TEXT,
                description TEXT,
                created_at TEXT NOT NULL
            );

            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                original_name TEXT NOT NULL,
                stored_name TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
                row_count INTEGER DEFAULT 0,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS materials (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id INTEGER NOT NULL,
                file_id INTEGER,
                sheet_name TEXT,
                excel_row INTEGER,
                item_code TEXT,
                item_name TEXT,
                specification TEXT,
                unit TEXT,
                quantity TEXT,
                supplier TEXT,
                po_no TEXT,
                request_date TEXT,
                required_date TEXT,
                status TEXT DEFAULT 'Chưa đặt hàng',
                location TEXT,
                responsible TEXT,
                actual_date TEXT,
                note TEXT,
                extra_json TEXT,
                updated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE,
                FOREIGN KEY(file_id) REFERENCES files(id) ON DELETE SET NULL
            );
            """
        )


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def detect_header(row_values: List[Any]) -> Tuple[Dict[str, int], int]:
    normalized = [normalize_text(v) for v in row_values]
    mapping: Dict[str, int] = {}
    score = 0
    for field, aliases in HEADER_ALIASES.items():
        aliases_norm = [normalize_text(a) for a in aliases]
        for idx, cell in enumerate(normalized):
            if not cell:
                continue
            if cell in aliases_norm or any(alias in cell for alias in aliases_norm):
                if field not in mapping:
                    mapping[field] = idx
                    score += 1
                break
    return mapping, score


def cell_to_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    return str(value).strip()


def parse_excel(path: Path) -> List[Dict[str, Any]]:
    workbook = load_workbook(path, data_only=True, read_only=True)
    rows: List[Dict[str, Any]] = []
    for sheet in workbook.worksheets:
        all_rows = list(sheet.iter_rows(values_only=True))
        if not all_rows:
            continue

        best_index = None
        best_mapping: Dict[str, int] = {}
        best_score = 0
        for idx, row in enumerate(all_rows[:20]):
            mapping, score = detect_header(list(row))
            if score > best_score:
                best_score = score
                best_index = idx
                best_mapping = mapping

        if best_index is None or best_score < 2:
            # Fallback: assume first non-empty row is header and import basic columns.
            for idx, row in enumerate(all_rows):
                if any(cell_to_text(v) for v in row):
                    best_index = idx
                    break
            if best_index is None:
                continue
            best_mapping = {
                "item_code": 0,
                "item_name": 1,
                "specification": 2,
                "unit": 3,
                "quantity": 4,
            }

        headers = [cell_to_text(v) or f"Column {i+1}" for i, v in enumerate(all_rows[best_index])]
        for row_number, row in enumerate(all_rows[best_index + 1 :], start=best_index + 2):
            if not any(cell_to_text(v) for v in row):
                continue
            item: Dict[str, Any] = {
                "sheet_name": sheet.title,
                "excel_row": row_number,
                "extra_json": {},
            }
            for field, col_index in best_mapping.items():
                item[field] = cell_to_text(row[col_index]) if col_index < len(row) else ""

            for col_index, header in enumerate(headers):
                value = cell_to_text(row[col_index]) if col_index < len(row) else ""
                if value:
                    item["extra_json"][header] = value

            if item.get("item_name") or item.get("item_code") or item.get("specification"):
                rows.append(item)
    workbook.close()
    return rows


@app.context_processor
def inject_globals():
    return {"STATUS_OPTIONS": STATUS_OPTIONS}


@app.route("/")
def index():
    with db() as conn:
        projects = conn.execute(
            """
            SELECT p.*,
                   COUNT(m.id) AS total_items,
                   SUM(CASE WHEN m.status = 'Đã về kho' OR m.status = 'Đã bàn giao công trình' THEN 1 ELSE 0 END) AS done_items,
                   SUM(CASE WHEN m.status = 'Thiếu / Chờ bổ sung' THEN 1 ELSE 0 END) AS issue_items
            FROM projects p
            LEFT JOIN materials m ON m.project_id = p.id
            GROUP BY p.id
            ORDER BY p.created_at DESC
            """
        ).fetchall()
    return render_template("index.html", projects=projects)


@app.route("/projects", methods=["POST"])
def create_project():
    name = request.form.get("name", "").strip()
    code = request.form.get("code", "").strip()
    description = request.form.get("description", "").strip()
    if not name:
        flash("Vui lòng nhập tên dự án.", "error")
        return redirect(url_for("index"))
    try:
        with db() as conn:
            conn.execute(
                "INSERT INTO projects(name, code, description, created_at) VALUES (?, ?, ?, ?)",
                (name, code, description, datetime.now().isoformat(timespec="seconds")),
            )
        flash("Đã tạo dự án.", "success")
    except sqlite3.IntegrityError:
        flash("Tên dự án đã tồn tại.", "error")
    return redirect(url_for("index"))


@app.route("/project/<int:project_id>")
def project_detail(project_id: int):
    status = request.args.get("status", "")
    keyword = request.args.get("q", "").strip()
    with db() as conn:
        project = conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
        if not project:
            flash("Không tìm thấy dự án.", "error")
            return redirect(url_for("index"))

        query = "SELECT * FROM materials WHERE project_id=?"
        params: List[Any] = [project_id]
        if status:
            query += " AND status=?"
            params.append(status)
        if keyword:
            query += " AND (item_name LIKE ? OR item_code LIKE ? OR specification LIKE ? OR po_no LIKE ? OR supplier LIKE ?)"
            like = f"%{keyword}%"
            params.extend([like, like, like, like, like])
        query += " ORDER BY id DESC"
        materials = conn.execute(query, params).fetchall()
        files = conn.execute("SELECT * FROM files WHERE project_id=? ORDER BY uploaded_at DESC", (project_id,)).fetchall()
        stats = conn.execute(
            "SELECT status, COUNT(*) AS count FROM materials WHERE project_id=? GROUP BY status",
            (project_id,),
        ).fetchall()
    return render_template("project.html", project=project, materials=materials, files=files, stats=stats, status=status, keyword=keyword)


@app.route("/project/<int:project_id>/upload", methods=["POST"])
def upload_excel(project_id: int):
    file = request.files.get("excel_file")
    if not file or not file.filename:
        flash("Vui lòng chọn file Excel.", "error")
        return redirect(url_for("project_detail", project_id=project_id))
    if not allowed_file(file.filename):
        flash("Chỉ hỗ trợ file .xlsx hoặc .xlsm.", "error")
        return redirect(url_for("project_detail", project_id=project_id))

    project_folder = UPLOAD_DIR / str(project_id)
    project_folder.mkdir(parents=True, exist_ok=True)
    original_name = secure_filename(file.filename)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    stored_name = f"{timestamp}_{original_name}"
    stored_path = project_folder / stored_name
    file.save(stored_path)

    try:
        rows = parse_excel(stored_path)
    except Exception as exc:
        flash(f"Không đọc được file Excel: {exc}", "error")
        return redirect(url_for("project_detail", project_id=project_id))

    with db() as conn:
        cur = conn.execute(
            "INSERT INTO files(project_id, original_name, stored_name, uploaded_at, row_count) VALUES (?, ?, ?, ?, ?)",
            (project_id, file.filename, stored_name, datetime.now().isoformat(timespec="seconds"), len(rows)),
        )
        file_id = cur.lastrowid
        for item in rows:
            conn.execute(
                """
                INSERT INTO materials(
                    project_id, file_id, sheet_name, excel_row, item_code, item_name, specification,
                    unit, quantity, supplier, po_no, request_date, required_date, status,
                    location, responsible, actual_date, note, extra_json, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    project_id,
                    file_id,
                    item.get("sheet_name"),
                    item.get("excel_row"),
                    item.get("item_code", ""),
                    item.get("item_name", ""),
                    item.get("specification", ""),
                    item.get("unit", ""),
                    item.get("quantity", ""),
                    item.get("supplier", ""),
                    item.get("po_no", ""),
                    item.get("request_date", ""),
                    item.get("required_date", ""),
                    "Chưa đặt hàng",
                    "",
                    "",
                    "",
                    item.get("note", ""),
                    json.dumps(item.get("extra_json", {}), ensure_ascii=False),
                    datetime.now().isoformat(timespec="seconds"),
                ),
            )
    flash(f"Đã upload và import {len(rows)} dòng vật tư.", "success")
    return redirect(url_for("project_detail", project_id=project_id))


@app.route("/material/<int:material_id>/update", methods=["POST"])
def update_material(material_id: int):
    status = request.form.get("status", "Chưa đặt hàng")
    if status not in STATUS_OPTIONS:
        status = "Chưa đặt hàng"
    fields = {
        "status": status,
        "location": request.form.get("location", "").strip(),
        "responsible": request.form.get("responsible", "").strip(),
        "actual_date": request.form.get("actual_date", "").strip(),
        "note": request.form.get("note", "").strip(),
        "updated_at": datetime.now().isoformat(timespec="seconds"),
    }
    with db() as conn:
        material = conn.execute("SELECT project_id FROM materials WHERE id=?", (material_id,)).fetchone()
        if not material:
            flash("Không tìm thấy vật tư.", "error")
            return redirect(url_for("index"))
        conn.execute(
            """
            UPDATE materials
            SET status=:status, location=:location, responsible=:responsible,
                actual_date=:actual_date, note=:note, updated_at=:updated_at
            WHERE id=:id
            """,
            {**fields, "id": material_id},
        )
    flash("Đã cập nhật tình trạng vật tư.", "success")
    return redirect(url_for("project_detail", project_id=material["project_id"]))


@app.route("/project/<int:project_id>/export.csv")
def export_csv(project_id: int):
    with db() as conn:
        project = conn.execute("SELECT * FROM projects WHERE id=?", (project_id,)).fetchone()
        materials = conn.execute("SELECT * FROM materials WHERE project_id=? ORDER BY id", (project_id,)).fetchall()
    if not project:
        return redirect(url_for("index"))

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow([
        "Mã vật tư", "Tên vật tư", "Quy cách", "Đơn vị", "Số lượng", "Nhà cung cấp", "PO",
        "Ngày yêu cầu", "Ngày cần", "Tình trạng", "Vị trí", "Người phụ trách", "Ngày thực tế", "Ghi chú"
    ])
    for m in materials:
        writer.writerow([
            m["item_code"], m["item_name"], m["specification"], m["unit"], m["quantity"], m["supplier"], m["po_no"],
            m["request_date"], m["required_date"], m["status"], m["location"], m["responsible"], m["actual_date"], m["note"]
        ])
    filename = f"{secure_filename(project['name'])}_materials.csv"
    return Response(
        output.getvalue().encode("utf-8-sig"),
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"},
    )


@app.route("/project/<int:project_id>/delete", methods=["POST"])
def delete_project(project_id: int):
    with db() as conn:
        conn.execute("DELETE FROM projects WHERE id=?", (project_id,))
    flash("Đã xóa dự án và dữ liệu liên quan.", "success")
    return redirect(url_for("index"))


if __name__ == "__main__":
    init_db()
    port = int(os.environ.get("PORT", 5000))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=port, debug=debug)
