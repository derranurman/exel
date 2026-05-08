#!/usr/bin/env python3
"""
Builder untuk file Sistem_Keluar_Masuk_Barang.xlsx
Dibuat pakai Python stdlib saja (tanpa openpyxl) karena sandbox tidak punya koneksi pip.

Struktur xlsx = zip berisi XML. Kita bangun manual.

Fitur yang di-embed ke dalam file:
 1. 4 Sheet: Input, Output Agus, Output Rexa, Rekap
 2. Header baris 1 (bold) + lebar kolom wajar.
 3. Data Validation:
    - Sheet Input, kolom B (Kode Barang): custom formula anti-duplikat
      COUNTIF($B$2:$B$1000,B2)<=1  -> error "Kode sudah ada"
    - Sheet Output Agus & Output Rexa, kolom B: list = Input!$B$2:$B$1000
      -> error "Kode tidak ada di Input"
 4. Formula auto tanggal di kolom C (butuh 'Iterative Calculation' ON di Excel):
      =IF(B2="","",IF(C2="",NOW(),C2))
      Formatnya DateTime indonesia.
 5. Sheet Rekap berisi:
      Total Input, Total Output Agus, Total Output Rexa, Total Output, Selisih
      Rekap per kode (Kode | Input | Agus | Rexa | Sisa).
 6. Conditional formatting highlight duplikat di Input (merah).
"""

import os
import zipfile
from xml.sax.saxutils import escape

OUT_PATH = os.path.join(os.path.dirname(__file__), "Sistem_Keluar_Masuk_Barang.xlsx")

# --------- Definisi sheet ----------
MAX_ROWS = 1000  # baris data max

SHEETS = [
    {
        "name": "Input Barang",
        "headers": ["No", "Kode Barang", "Tanggal Masuk"],
        "role": "input",
    },
    {
        "name": "Output Agus",
        "headers": ["No", "Kode Barang", "Tanggal Keluar"],
        "role": "output",
    },
    {
        "name": "Output Rexa",
        "headers": ["No", "Kode Barang", "Tanggal Keluar"],
        "role": "output",
    },
    {
        "name": "Rekap",
        "headers": [],  # diisi manual di bawah
        "role": "rekap",
    },
]


def col_letter(idx):
    """1 -> A, 2 -> B, ..."""
    s = ""
    while idx > 0:
        idx, r = divmod(idx - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------- [Content_Types].xml ----------
def content_types_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    ]
    for i, _ in enumerate(SHEETS, start=1):
        parts.append(
            f'<Override PartName="/xl/worksheets/sheet{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        )
    parts.append(
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    )
    parts.append(
        '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
    )
    parts.append(
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
    )
    parts.append(
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
    )
    parts.append("</Types>")
    return "".join(parts)


# ---------- _rels/.rels ----------
def root_rels_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        "</Relationships>"
    )


# ---------- docProps ----------
def core_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        "<dc:creator>Kiro</dc:creator>"
        "<dc:title>Sistem Keluar Masuk Barang</dc:title>"
        "</cp:coreProperties>"
    )


def app_xml():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        "<Application>Kiro</Application>"
        "</Properties>"
    )


# ---------- xl/workbook.xml ----------
def workbook_xml():
    sheets_xml = "".join(
        f'<sheet name="{escape(s["name"])}" sheetId="{i}" r:id="rId{i}"/>'
        for i, s in enumerate(SHEETS, start=1)
    )
    # calcPr iterateCount=100 supaya auto-tanggal formula bekerja
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<workbookPr defaultThemeVersion="166925"/>'
        f"<sheets>{sheets_xml}</sheets>"
        '<calcPr calcId="162913" iterate="1" iterateCount="100" iterateDelta="0.001"/>'
        "</workbook>"
    )


def workbook_rels_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i, _ in enumerate(SHEETS, start=1):
        parts.append(
            f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>'
        )
    # styles & sharedStrings
    parts.append(
        f'<Relationship Id="rId{len(SHEETS)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    parts.append(
        f'<Relationship Id="rId{len(SHEETS)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
    )
    parts.append("</Relationships>")
    return "".join(parts)


# ---------- xl/styles.xml ----------
# Style index:
# 0 = default
# 1 = bold header (fill abu-abu, border, center)
# 2 = date-time format (dd/mm/yyyy hh:mm)
# 3 = bold big label (rekap title)
# 4 = number with border
# 5 = text with border
# 6 = bold with border (rekap header)
def styles_xml():
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1">
  <numFmt numFmtId="164" formatCode="dd/mm/yyyy hh:mm"/>
</numFmts>
<fonts count="3">
  <font><sz val="11"/><name val="Calibri"/></font>
  <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
  <font><b/><sz val="14"/><name val="Calibri"/></font>
</fonts>
<fills count="3">
  <fill><patternFill patternType="none"/></fill>
  <fill><patternFill patternType="gray125"/></fill>
  <fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="2">
  <border><left/><right/><top/><bottom/><diagonal/></border>
  <border>
    <left style="thin"><color rgb="FF808080"/></left>
    <right style="thin"><color rgb="FF808080"/></right>
    <top style="thin"><color rgb="FF808080"/></top>
    <bottom style="thin"><color rgb="FF808080"/></bottom>
    <diagonal/>
  </border>
</borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="7">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
  <xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
  <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>
  <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
<dxfs count="1">
  <dxf><font><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
</dxfs>
<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>
</styleSheet>
"""


# ---------- sharedStrings (kita simpan semua string unique) ----------
SHARED = []
SHARED_INDEX = {}


def s(text):
    """Tambah string ke shared strings table, return indeksnya."""
    if text not in SHARED_INDEX:
        SHARED_INDEX[text] = len(SHARED)
        SHARED.append(text)
    return SHARED_INDEX[text]


def shared_strings_xml():
    parts = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{len(SHARED)}" uniqueCount="{len(SHARED)}">',
    ]
    for t in SHARED:
        parts.append(f"<si><t xml:space=\"preserve\">{escape(t)}</t></si>")
    parts.append("</sst>")
    return "".join(parts)


# ---------- Cell helpers ----------
def cell_str(col, row, text, style=0):
    idx = s(text)
    return f'<c r="{col}{row}" s="{style}" t="s"><v>{idx}</v></c>'


def cell_num(col, row, num, style=0):
    return f'<c r="{col}{row}" s="{style}"><v>{num}</v></c>'


def cell_formula(col, row, formula, style=0, cache=""):
    v = f"<v>{cache}</v>" if cache != "" else ""
    return f'<c r="{col}{row}" s="{style}"><f>{escape(formula)}</f>{v}</c>'


def cell_blank(col, row, style=0):
    return f'<c r="{col}{row}" s="{style}"/>'


# ---------- Worksheet builders ----------
def build_data_sheet(s_def, is_input):
    """Sheet Input Barang / Output Agus / Output Rexa.

    Kolom: A=No (formula), B=Kode Barang (text), C=Tanggal (formula NOW)
    """
    name = s_def["name"]

    rows_xml = []

    # --- Row 1 : header ---
    header_cells = []
    for i, h in enumerate(s_def["headers"], start=1):
        header_cells.append(cell_str(col_letter(i), 1, h, style=1))
    rows_xml.append(f'<row r="1" ht="22" customHeight="1">{"".join(header_cells)}</row>')

    # --- Row 2..MAX : data kosong dengan formula pre-filled ---
    for r in range(2, MAX_ROWS + 1):
        # Kolom A: Nomor otomatis -> =IF(B2="","",ROW()-1)
        a = cell_formula("A", r, f'IF(B{r}="","",ROW()-1)', style=4)
        # Kolom B: tempat scan kode
        b = cell_blank("B", r, style=5)
        # Kolom C: auto tanggal -> =IF(B2="","",IF(C2="",NOW(),C2))
        c_formula = f'IF(B{r}="","",IF(C{r}="",NOW(),C{r}))'
        c = cell_formula("C", r, c_formula, style=2)
        rows_xml.append(f'<row r="{r}">{a}{b}{c}</row>')

    # --- dataValidations ---
    dv_xml = ""
    if is_input:
        # anti duplikat
        formula = f'COUNTIF($B$2:$B${MAX_ROWS},B2)=1'
        dv_xml = (
            f'<dataValidations count="1">'
            f'<dataValidation type="custom" allowBlank="1" showInputMessage="0" '
            f'showErrorMessage="1" errorStyle="stop" '
            f'errorTitle="Kode Duplikat" '
            f'error="Kode barang ini SUDAH ADA di Input. Gunakan kode lain." '
            f'sqref="B2:B{MAX_ROWS}">'
            f'<formula1>{escape(formula)}</formula1>'
            f'</dataValidation>'
            f'</dataValidations>'
        )
    else:
        # Validasi: kode HARUS sudah ada di Input Barang.
        # Pakai type=custom (bukan list) supaya scanner gak ketemu dropdown
        # dan langsung bisa tekan Enter.
        formula = (
            f"COUNTIF('Input Barang'!$B$2:$B${MAX_ROWS},B2)>=1"
        )
        dv_xml = (
            f'<dataValidations count="1">'
            f'<dataValidation type="custom" allowBlank="1" showInputMessage="0" '
            f'showErrorMessage="1" errorStyle="stop" '
            f'errorTitle="Kode Tidak Terdaftar" '
            f'error="Kode barang ini TIDAK ADA di sheet Input Barang. Input dulu di sheet Input Barang." '
            f'sqref="B2:B{MAX_ROWS}">'
            f'<formula1>{escape(formula)}</formula1>'
            f'</dataValidation>'
            f'</dataValidations>'
        )

    # --- conditional formatting (highlight duplikat di input) ---
    cf_xml = ""
    if is_input:
        cf_formula = f'COUNTIF($B$2:$B${MAX_ROWS},B2)>1'
        cf_xml = (
            f'<conditionalFormatting sqref="B2:B{MAX_ROWS}">'
            f'<cfRule type="expression" dxfId="0" priority="1">'
            f'<formula>{escape(cf_formula)}</formula>'
            f'</cfRule>'
            f'</conditionalFormatting>'
        )

    # --- column widths + freeze header + sheetView activeCell B2 ---
    cols = (
        '<cols>'
        '<col min="1" max="1" width="6" customWidth="1"/>'
        '<col min="2" max="2" width="22" customWidth="1"/>'
        '<col min="3" max="3" width="22" customWidth="1"/>'
        '</cols>'
    )

    sheet_views = (
        '<sheetViews>'
        '<sheetView workbookViewId="0" tabSelected="1">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="B2" sqref="B2"/>'
        '</sheetView>'
        '</sheetViews>'
    )

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<dimension ref="A1:C' + str(MAX_ROWS) + '"/>'
        + sheet_views
        + '<sheetFormatPr defaultRowHeight="15"/>'
        + cols
        + '<sheetData>' + "".join(rows_xml) + '</sheetData>'
        # IMPORTANT: per OOXML schema, conditionalFormatting MUST come
        # before dataValidations. Kalau terbalik Excel menolak loadnya.
        + cf_xml
        + dv_xml
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + '</worksheet>'
    )
    return xml


def build_rekap_sheet():
    MAXR = MAX_ROWS

    rows_xml = []

    # --- Bagian atas: total ---
    # A1 = "REKAP KELUAR MASUK BARANG" (style 3 big bold)
    rows_xml.append(
        f'<row r="1" ht="26" customHeight="1">'
        f'{cell_str("A", 1, "REKAP KELUAR MASUK BARANG", style=3)}'
        f'</row>'
    )

    # Row 3 header ringkasan
    rows_xml.append(
        f'<row r="3">'
        f'{cell_str("A", 3, "Keterangan", style=6)}'
        f'{cell_str("B", 3, "Jumlah", style=6)}'
        f'</row>'
    )
    # Row 4..8 data
    f_input = "COUNTA('Input Barang'!B2:B" + str(MAXR) + ")"
    f_agus = "COUNTA('Output Agus'!B2:B" + str(MAXR) + ")"
    f_rexa = "COUNTA('Output Rexa'!B2:B" + str(MAXR) + ")"
    rows_xml.append(
        f'<row r="4">'
        f'{cell_str("A", 4, "Total Barang Masuk (Input)", style=5)}'
        f'{cell_formula("B", 4, f_input, style=4)}'
        f'</row>'
    )
    rows_xml.append(
        f'<row r="5">'
        f'{cell_str("A", 5, "Total Keluar - Agus", style=5)}'
        f'{cell_formula("B", 5, f_agus, style=4)}'
        f'</row>'
    )
    rows_xml.append(
        f'<row r="6">'
        f'{cell_str("A", 6, "Total Keluar - Rexa", style=5)}'
        f'{cell_formula("B", 6, f_rexa, style=4)}'
        f'</row>'
    )
    rows_xml.append(
        f'<row r="7">'
        f'{cell_str("A", 7, "Total Keluar (Agus + Rexa)", style=5)}'
        f'{cell_formula("B", 7, "B5+B6", style=4)}'
        f'</row>'
    )
    rows_xml.append(
        f'<row r="8">'
        f'{cell_str("A", 8, "Sisa Stok (Input - Keluar)", style=5)}'
        f'{cell_formula("B", 8, "B4-B7", style=4)}'
        f'</row>'
    )

    # --- Bagian bawah: Rekap per kode ---
    rows_xml.append(
        f'<row r="10" ht="20" customHeight="1">'
        f'{cell_str("A", 10, "REKAP PER KODE BARANG", style=3)}'
        f'</row>'
    )
    rows_xml.append(
        f'<row r="11">'
        f'{cell_str("A", 11, "Kode Barang", style=6)}'
        f'{cell_str("B", 11, "Input", style=6)}'
        f'{cell_str("C", 11, "Keluar Agus", style=6)}'
        f'{cell_str("D", 11, "Keluar Rexa", style=6)}'
        f'{cell_str("E", 11, "Sisa", style=6)}'
        f'</row>'
    )

    # Rows 12..MAXR: rekap per kode
    # A12 = IF(ROW()-11<=COUNTA(Input!B:B)-0, INDEX(Input!B:B, ROW()-10), "")
    # Lebih simpel: =IFERROR(INDEX('Input Barang'!$B$2:$B$1000, ROW()-11), "")
    for rr in range(12, MAXR + 1):
        relative = rr - 11  # 1,2,3,...
        f_a = (
            "IFERROR(INDEX('Input Barang'!$B$2:$B$" + str(MAXR) + ","
            + str(relative) + '),"")'
        )
        f_b = (
            'IF(A' + str(rr) + '="","",'
            "COUNTIF('Input Barang'!$B$2:$B$" + str(MAXR) + ",A" + str(rr) + "))"
        )
        f_c = (
            'IF(A' + str(rr) + '="","",'
            "COUNTIF('Output Agus'!$B$2:$B$" + str(MAXR) + ",A" + str(rr) + "))"
        )
        f_d = (
            'IF(A' + str(rr) + '="","",'
            "COUNTIF('Output Rexa'!$B$2:$B$" + str(MAXR) + ",A" + str(rr) + "))"
        )
        f_e = (
            'IF(A' + str(rr) + '="","",B' + str(rr) + "-C" + str(rr)
            + "-D" + str(rr) + ")"
        )
        a = cell_formula("A", rr, f_a, style=5)
        b = cell_formula("B", rr, f_b, style=4)
        c = cell_formula("C", rr, f_c, style=4)
        d = cell_formula("D", rr, f_d, style=4)
        e = cell_formula("E", rr, f_e, style=4)
        rows_xml.append(f'<row r="{rr}">{a}{b}{c}{d}{e}</row>')

    cols = (
        '<cols>'
        '<col min="1" max="1" width="30" customWidth="1"/>'
        '<col min="2" max="5" width="16" customWidth="1"/>'
        '</cols>'
    )

    sheet_views = (
        '<sheetViews>'
        '<sheetView workbookViewId="0">'
        '<selection activeCell="A1" sqref="A1"/>'
        '</sheetView>'
        '</sheetViews>'
    )

    xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        '<dimension ref="A1:E' + str(MAX_ROWS) + '"/>'
        + sheet_views
        + '<sheetFormatPr defaultRowHeight="15"/>'
        + cols
        + '<sheetData>' + "".join(rows_xml) + '</sheetData>'
        + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + '</worksheet>'
    )
    return xml


# ---------- Main ----------
def main():
    sheet_xmls = []
    for sdef in SHEETS:
        role = sdef["role"]
        if role == "input":
            sheet_xmls.append(build_data_sheet(sdef, is_input=True))
        elif role == "output":
            sheet_xmls.append(build_data_sheet(sdef, is_input=False))
        elif role == "rekap":
            sheet_xmls.append(build_rekap_sheet())

    # shared strings harus dibuat SETELAH semua sheet dibuild (karena string
    # ditambah selama build lewat s()).
    ss_xml = shared_strings_xml()

    if os.path.exists(OUT_PATH):
        os.remove(OUT_PATH)

    with zipfile.ZipFile(OUT_PATH, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types_xml())
        z.writestr("_rels/.rels", root_rels_xml())
        z.writestr("docProps/core.xml", core_xml())
        z.writestr("docProps/app.xml", app_xml())
        z.writestr("xl/workbook.xml", workbook_xml())
        z.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml())
        z.writestr("xl/styles.xml", styles_xml())
        z.writestr("xl/sharedStrings.xml", ss_xml)
        for i, content in enumerate(sheet_xmls, start=1):
            z.writestr(f"xl/worksheets/sheet{i}.xml", content)

    print(f"OK -> {OUT_PATH}")
    print(f"Size: {os.path.getsize(OUT_PATH)} bytes")


if __name__ == "__main__":
    main()
