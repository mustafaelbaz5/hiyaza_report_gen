from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .models import ReportInput


def _clean(value: str) -> str:
    return str(value or "").strip()


def _find_value_cell(table, label: str):
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            if _clean(cell.text) == label and index + 1 < len(row.cells):
                return row.cells[index + 1]
    return None


def _set_cell_text(cell, value: str):
    if cell is None:
        return
    # Preserve the formatting of the first available run.
    first_run = None
    for paragraph in cell.paragraphs:
        if paragraph.runs:
            first_run = paragraph.runs[0]
            break
    font = None
    if first_run is not None:
        font = {
            "name": first_run.font.name,
            "size": first_run.font.size,
            "bold": first_run.bold,
            "italic": first_run.italic,
            "color": first_run.font.color.rgb,
        }
    cell.text = str(value)
    paragraph = cell.paragraphs[0]
    paragraph.alignment = 1
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)
    bidi.set(qn("w:val"), "1")
    if paragraph.runs and font:
        run = paragraph.runs[0]
        run.font.name = font["name"]
        run.font.size = font["size"]
        run.bold = font["bold"]
        run.italic = font["italic"]
        if font["color"] is not None:
            run.font.color.rgb = font["color"]
        r_pr = run._r.get_or_add_rPr()
        rtl = r_pr.find(qn("w:rtl"))
        if rtl is None:
            rtl = OxmlElement("w:rtl")
            r_pr.append(rtl)
        rtl.set(qn("w:val"), "1")


def _find_row(table, label: str):
    for row in table.rows:
        if any(_clean(cell.text) == label for cell in row.cells):
            return row
    return None


def validate_template(doc: Document):
    required_labels = {
        "الجهة المستلمة", "تاريخ التسليم", "الجهة المنفذة", "حالة التقرير",
        "اسم الجمعية الرسمي", "نوع الجمعية", "كود الجمعية", "عدد الأحواض",
        "إجمالي عدد القطع", "القطع المسجلة على المنظومة",
        "القطع غير المسجلة على المنظومة", "منها قطع غير محيزة",
    }
    found = set()
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = _clean(cell.text)
                if text in required_labels:
                    found.add(text)
    missing = sorted(required_labels - found)
    if missing:
        raise ValueError("قالب Word غير متوافق. الحقول الناقصة: " + "، ".join(missing))
    if len(doc.tables) < 5:
        raise ValueError("قالب Word غير متوافق: جدول الملاحظات أو المرفقات غير موجود.")


def generate_report(template_path: str | Path, output_path: str | Path, data: ReportInput):
    template = Path(template_path)
    output = Path(output_path)
    if not template.exists():
        raise FileNotFoundError(f"ملف القالب غير موجود: {template}")

    doc = Document(template)
    validate_template(doc)

    replacements = {
        "الجهة المستلمة": data.recipient,
        "تاريخ التسليم": data.delivery_date,
        "الجهة المنفذة": data.executor,
        "حالة التقرير": data.report_status,
        "اسم الجمعية الرسمي": data.association_name,
        "نوع الجمعية": data.association_type,
        "كود الجمعية": data.association_code,
        "عدد الأحواض": str(data.analysis.basin_count),
    }
    for table in doc.tables:
        for label, value in replacements.items():
            cell = _find_value_cell(table, label)
            if cell is not None:
                _set_cell_text(cell, value)

    stats = {
        "إجمالي عدد القطع": (data.analysis.total_parcels, 100.0),
        "القطع المسجلة على المنظومة": (data.analysis.registered_count, data.analysis.registered_percentage),
        "القطع غير المسجلة على المنظومة": (data.analysis.unregistered_count, data.analysis.unregistered_percentage),
        "منها قطع غير محيزة": (data.analysis.non_held_count, data.analysis.non_held_percentage),
    }
    for table in doc.tables:
        for label, (count, percentage) in stats.items():
            row = _find_row(table, label)
            if row is None or len(row.cells) < 3:
                continue
            # Template visual order: label, count, percentage in XML order.
            _set_cell_text(row.cells[1], str(count))
            percentage_text = "100%" if percentage == 100 else f"{percentage:.1f}%"
            _set_cell_text(row.cells[2], percentage_text)

    # Notes table is identified by its header.
    for table in doc.tables:
        if any("ملخص الحالات الخاصة" in _clean(c.text) for r in table.rows for c in r.cells):
            if len(table.rows) >= 2:
                _set_cell_text(table.cell(1, 0), data.special_notes or "لا يوجد")
            break

    package_name = f"ملف التسليم المضغوط باسم {data.village_name}"
    for table in doc.tables:
        if any("حزمة التسليم المضغوطة" in _clean(c.text) for r in table.rows for c in r.cells):
            if len(table.rows) >= 2:
                _set_cell_text(
                    table.cell(1, 0),
                    package_name + "\nيحتوي على قاعدة البيانات الجغرافية وملف بيانات القطع Excel",
                )
            break

    output.parent.mkdir(parents=True, exist_ok=True)
    doc.core_properties.title = f"تقرير تسليم {data.association_name}"
    doc.save(output)
    return output
