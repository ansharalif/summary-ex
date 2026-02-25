import os
from datetime import datetime
from typing import List, Dict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


# =========================================================
# KONFIGURASI DATA CONTOH (3-5 negara terisi)
# =========================================================
SAMPLE_DATA: List[Dict[str, str]] = [
    {
        "Negara": "Indonesia",
        "Last update": "2026-02-20",
        "Pendekatan dominan": "Kombinasi penegakan hukum, pencegahan, dan rehabilitasi",
        "Instrumen hukum utama": "Peraturan perundang-undangan terkait tindak pidana terorisme dan kebijakan turunan",
        "Catatan kehati-hatian (HAM/rule of law)": "Perlu menjaga keseimbangan antara keamanan, due process, dan perlindungan hak-hak dasar",
        "Relevansi pembelajaran untuk Indonesia/BNPT": "Penguatan standardisasi pemetaan kebijakan dan monitoring pembaruan lintas isu",
        "Sumber ringkas": "https://www.foreignterroristfighters.info/policy-matrix",
    },
    {
        "Negara": "France",
        "Last update": "2026-02-15",
        "Pendekatan dominan": "Penegakan hukum dan pengawasan yang kuat, disertai langkah pencegahan",
        "Instrumen hukum utama": "Kerangka hukum nasional kontra-terorisme dan mekanisme proses pidana",
        "Catatan kehati-hatian (HAM/rule of law)": "Perlu perhatian pada proporsionalitas pengawasan dan perlindungan kebebasan sipil",
        "Relevansi pembelajaran untuk Indonesia/BNPT": "Pembelajaran komparatif untuk aspek koordinasi penegakan dan pengawasan berbasis hukum",
        "Sumber ringkas": "https://www.foreignterroristfighters.info/policy-matrix",
    },
    {
        "Negara": "Germany",
        "Last update": "2026-02-12",
        "Pendekatan dominan": "Kombinasi penegakan, pencegahan, dan program rehabilitasi/reintegrasi",
        "Instrumen hukum utama": "Ketentuan pidana, pengawasan, dan kebijakan reintegrasi berbasis kelembagaan",
        "Catatan kehati-hatian (HAM/rule of law)": "Penting memastikan non-diskriminasi, due process, dan evaluasi efektivitas program",
        "Relevansi pembelajaran untuk Indonesia/BNPT": "Referensi untuk pengembangan kerangka pembelajaran rehabilitasi dan koordinasi multi-aktor",
        "Sumber ringkas": "https://www.foreignterroristfighters.info/policy-matrix",
    },
    {
        "Negara": "Kazakhstan",
        "Last update": "2026-02-05",
        "Pendekatan dominan": "Kombinasi repatriasi, reintegrasi, dan langkah keamanan",
        "Instrumen hukum utama": "Kebijakan nasional penanganan returnees dan dukungan kelembagaan terkait",
        "Catatan kehati-hatian (HAM/rule of law)": "Perlu perhatian pada perlindungan kelompok rentan dan akuntabilitas implementasi",
        "Relevansi pembelajaran untuk Indonesia/BNPT": "Bahan belajar untuk isu repatriasi, reintegrasi, dan dukungan keluarga/perempuan-anak",
        "Sumber ringkas": "https://www.foreignterroristfighters.info/policy-matrix",
    },
    {
        "Negara": "Philippines",
        "Last update": "2026-01-28",
        "Pendekatan dominan": "Penegakan hukum dengan dukungan pencegahan; data reintegrasi perlu diperdalam",
        "Instrumen hukum utama": "Perangkat hukum kontra-terorisme nasional dan kebijakan operasional terkait",
        "Catatan kehati-hatian (HAM/rule of law)": "Perlu menjaga akuntabilitas, pengawasan, dan perlindungan HAM dalam implementasi",
        "Relevansi pembelajaran untuk Indonesia/BNPT": "Pembelajaran regional untuk komparasi pendekatan dan identifikasi gap data kebijakan",
        "Sumber ringkas": "https://www.foreignterroristfighters.info/policy-matrix",
    },
]

COLUMNS = [
    "Negara",
    "Last update",
    "Pendekatan dominan",
    "Instrumen hukum utama",
    "Catatan kehati-hatian (HAM/rule of law)",
    "Relevansi pembelajaran untuk Indonesia/BNPT",
    "Sumber ringkas",
]


# =========================================================
# UTIL STYLE EXCEL
# =========================================================
def apply_thin_border(cell):
    thin = Side(style="thin", color="D9D9D9")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def style_header_row(ws, row_idx: int, start_col: int, end_col: int):
    fill = PatternFill(fill_type="solid", fgColor="1F4E78")  # biru tua
    font = Font(color="FFFFFF", bold=True)
    align = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = align
        apply_thin_border(cell)


def auto_width(ws, min_width: int = 12, max_width: int = 55):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for c in col_cells:
            val = "" if c.value is None else str(c.value)
            max_len = max(max_len, len(val))
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


# =========================================================
# EXCEL GENERATOR
# =========================================================
def create_excel_template(output_path: str, sample_data: List[Dict[str, str]]):
    wb = Workbook()

    # Sheet 1: TEMPLATE
    ws_template = wb.active
    ws_template.title = "Template"

    ws_template["A1"] = "Template Policy Matrix Ringkas - Analisis Cepat Pimpinan"
    ws_template["A1"].font = Font(size=13, bold=True, color="1F1F1F")
    ws_template.merge_cells("A1:G1")

    ws_template["A2"] = "Isi 1 baris = 1 country brief ringkas (1 halaman saat diekspor ke Word)"
    ws_template["A2"].font = Font(italic=True, color="666666")
    ws_template.merge_cells("A2:G2")

    # Header
    for idx, col_name in enumerate(COLUMNS, start=1):
        ws_template.cell(row=4, column=idx, value=col_name)
    style_header_row(ws_template, row_idx=4, start_col=1, end_col=len(COLUMNS))

    # 5 baris kosong template
    for r in range(5, 10):
        for c in range(1, len(COLUMNS) + 1):
            cell = ws_template.cell(row=r, column=c, value="" if c != 2 else "YYYY-MM-DD")
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            apply_thin_border(cell)

    # Freeze pane
    ws_template.freeze_panes = "A5"

    # Sheet 2: CONTOH TERISI
    ws_sample = wb.create_sheet("Contoh_3-5_Negara")
    ws_sample["A1"] = "Contoh Isian Policy Matrix Ringkas (5 Negara)"
    ws_sample["A1"].font = Font(size=13, bold=True)
    ws_sample.merge_cells("A1:G1")

    for idx, col_name in enumerate(COLUMNS, start=1):
        ws_sample.cell(row=3, column=idx, value=col_name)
    style_header_row(ws_sample, row_idx=3, start_col=1, end_col=len(COLUMNS))

    for row_idx, item in enumerate(sample_data, start=4):
        for col_idx, col_name in enumerate(COLUMNS, start=1):
            cell = ws_sample.cell(row=row_idx, column=col_idx, value=item.get(col_name, ""))
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            apply_thin_border(cell)

    ws_sample.freeze_panes = "A4"

    # Sheet 3: PANDUAN
    ws_guide = wb.create_sheet("Panduan_Pengisian")
    ws_guide["A1"] = "Panduan Pengisian Singkat - Policy Matrix Ringkas"
    ws_guide["A1"].font = Font(size=13, bold=True)
    ws_guide.merge_cells("A1:D1")

    guide_rows = [
        ["Kolom", "Isi", "Contoh", "Catatan"],
        ["Negara", "Nama negara", "France", "1 baris untuk 1 negara / 1 isu"],
        ["Last update", "Tanggal pembaruan terakhir", "2026-02-15", "Gunakan format YYYY-MM-DD"],
        ["Pendekatan dominan", "Ringkasan pendekatan utama", "Penegakan + pencegahan", "Tulis singkat, 1-2 kalimat"],
        ["Instrumen hukum utama", "Instrumen/kerangka hukum kunci", "Ketentuan pidana + kebijakan turunan", "Tidak perlu terlalu detail"],
        ["Catatan kehati-hatian (HAM/rule of law)", "Poin kehati-hatian implementasi", "Due process, proporsionalitas", "Gunakan bahasa netral dan analitis"],
        ["Relevansi pembelajaran untuk Indonesia/BNPT", "Nilai pembelajaran untuk internal", "Komparasi, lesson learned", "Fokus pada pembelajaran, bukan evaluasi normatif"],
        ["Sumber ringkas", "URL sumber utama", "https://www.foreignterroristfighters.info/policy-matrix", "Cantumkan sumber untuk penelusuran"],
    ]

    for r_idx, row in enumerate(guide_rows, start=3):
        for c_idx, value in enumerate(row, start=1):
            cell = ws_guide.cell(row=r_idx, column=c_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            apply_thin_border(cell)

    style_header_row(ws_guide, row_idx=3, start_col=1, end_col=4)

    # Style umum
    for ws in [ws_template, ws_sample, ws_guide]:
        auto_width(ws)
        # Tinggi baris agar teks lebih terbaca
        for r in range(1, ws.max_row + 1):
            ws.row_dimensions[r].height = 22

    # Sedikit penyesuaian lebar kolom yang panjang
    for ws in [ws_template, ws_sample]:
        ws.column_dimensions["C"].width = 38
        ws.column_dimensions["D"].width = 38
        ws.column_dimensions["E"].width = 45
        ws.column_dimensions["F"].width = 45
        ws.column_dimensions["G"].width = 42

    wb.save(output_path)


# =========================================================
# DOCX UTIL
# =========================================================
def set_cell_shading(cell, fill: str):
    """fill hex string, e.g. 'D9E1F2'"""
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_doc_margins(document: Document, top=0.6, bottom=0.6, left=0.7, right=0.7):
    section = document.sections[0]
    section.top_margin = Inches(top)
    section.bottom_margin = Inches(bottom)
    section.left_margin = Inches(left)
    section.right_margin = Inches(right)


# =========================================================
# WORD GENERATOR (1 halaman per negara)
# =========================================================
def create_country_brief_docx(item: Dict[str, str], output_path: str):
    doc = Document()
    set_doc_margins(doc)

    # Judul
    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("POLICY MATRIX RINGKAS - ANALISIS CEPAT PIMPINAN")
    run.bold = True
    run.font.size = Pt(13)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run(f"Negara: {item.get('Negara', '-')}")
    run_sub.italic = True
    run_sub.font.size = Pt(10)

    doc.add_paragraph()  # spacer

    # Tabel 2 kolom
    table = doc.add_table(rows=0, cols=2)
    table.style = "Table Grid"
    table.autofit = False

    rows_content = [
        ("Negara", item.get("Negara", "-")),
        ("Last update", item.get("Last update", "-")),
        ("Pendekatan dominan", item.get("Pendekatan dominan", "-")),
        ("Instrumen hukum utama", item.get("Instrumen hukum utama", "-")),
        ("Catatan kehati-hatian (HAM/rule of law)", item.get("Catatan kehati-hatian (HAM/rule of law)", "-")),
        ("Relevansi pembelajaran untuk Indonesia/BNPT", item.get("Relevansi pembelajaran untuk Indonesia/BNPT", "-")),
        ("Sumber ringkas", item.get("Sumber ringkas", "-")),
    ]

    for label, value in rows_content:
        row_cells = table.add_row().cells
        row_cells[0].text = label
        row_cells[1].text = value

        # Lebar kolom
        row_cells[0].width = Inches(2.2)
        row_cells[1].width = Inches(4.9)

        # style text
        for i, c in enumerate(row_cells):
            for p in c.paragraphs:
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.space_before = Pt(1)
                for r in p.runs:
                    r.font.size = Pt(10)
                    if i == 0:
                        r.bold = True
            if i == 0:
                set_cell_shading(c, "D9E1F2")

    doc.add_paragraph()

    foot = doc.add_paragraph()
    foot_run = foot.add_run(
        "Catatan: Dokumen ini merupakan ringkasan internal untuk kebutuhan analisis cepat. "
        "Lakukan verifikasi lanjutan pada sumber primer sebelum digunakan untuk keputusan substantif."
    )
    foot_run.font.size = Pt(9)
    foot_run.italic = True

    doc.save(output_path)


# =========================================================
# MAIN PROGRAM
# =========================================================
def main():
    output_dir = "output_policy_matrix_ringkas"
    os.makedirs(output_dir, exist_ok=True)

    # 1) Buat Excel template + contoh + panduan
    excel_path = os.path.join(output_dir, "Template_Policy_Matrix_Ringkas.xlsx")
    create_excel_template(excel_path, SAMPLE_DATA)

    # 2) Buat Word 1 halaman per negara (5 contoh)
    generated_docs = []
    for item in SAMPLE_DATA:
        country_safe = item["Negara"].replace(" ", "_")
        docx_path = os.path.join(output_dir, f"Policy_Matrix_Ringkas_{country_safe}.docx")
        create_country_brief_docx(item, docx_path)
        generated_docs.append(docx_path)

    # 3) Log ringkas
    print("=" * 60)
    print("SELESAI - Generator Policy Matrix Ringkas")
    print("=" * 60)
    print(f"Waktu proses : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"Folder output : {os.path.abspath(output_dir)}")
    print(f"Excel template : {excel_path}")
    print("Word briefs    :")
    for p in generated_docs:
        print(f" - {p}")
    print("=" * 60)


if __name__ == "__main__":
    main()
