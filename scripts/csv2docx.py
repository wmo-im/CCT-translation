#!/usr/bin/env python3
"""
Generate Word (.docx) tables from CCT CSV files.

Usage:
    python scripts/csv2docx.py <lang> [--table C00|C01|...|all] [--outdir DIR]

Examples:
    python scripts/csv2docx.py ru                # all 13 tables
    python scripts/csv2docx.py fr --table C14     # French C14 only
    python scripts/csv2docx.py es --outdir /tmp   # Spanish, custom dir
"""

import argparse, csv, os, sys, glob

from docx import Document
from docx.shared import Pt, Cm
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

FONT = 'Arial'
SZ_TITLE = Pt(11)
SZ_HDR = Pt(8)
SZ_DATA = Pt(8)

# CCT table display names
TABLE_NAMES = {
    'C00': 'GRIB, BUFR and CREX master table version number',
    'C01': 'Identification of originating/generating centre',
    'C02': 'Radiosonde/sounding system used',
    'C03': 'Instrument make and type',
    'C04': 'Water temperature profile recorder types',
    'C05': 'Satellite identifier',
    'C06': 'List of units for TDCFs',
    'C07': 'Tracking techniques/status of system used',
    'C08': 'Satellite instruments',
    'C11': 'Originating/generating centres',
    'C12': 'Sub-centres of originating centres',
    'C13': 'Data sub-categories',
    'C14': 'Atmospheric chemical or physical constituent type',
}

# Which columns to show per table (beyond ID which is always first)
TABLE_COLS = {
    'C00': ['GRIB version number', 'BUFR version number', 'CREX version number', 'Effective date'],
    'C01': ['CodeFigureForF1F2', 'CodeFigureForF3F3F3', 'Octet5GRIB1_Octet6BUFR3', 'OriginatingGeneratingCentres_{lang}'],
    'C02': ['DateOfAssignment_{lang}', 'CodeFigureForrara', 'CodeFigureForBUFR', 'RadiosondeSoundingSystemUsed_{lang}'],
    'C03': ['CodeFigureForIXIIXIX', 'CodeFigureForBUFR', 'InstrumentMakeAndType_{lang}', 'EquationCoefficients_a', 'EquationCoefficients_b'],
    'C04': ['CodeFigureForXRXR', 'CodeFigureForBUFR', 'Meaning_{lang}'],
    'C05': ['CodeFigureForI6I6I6', 'CodeFigureForBUFR', 'CodeFigureForGRIB2', 'SatelliteName_{lang}'],
    'C06': ['CodeFigure', 'UnitType_{lang}', 'Meaning_{lang}', 'conventional', 'IA5-ASCII', 'ITA2', 'SIDefinition', 'Note', 'NoteID'],
    'C07': ['CodeFigureForsasa', 'CodeFigureForBUFR', 'TrackingTechniquesStatusOfSystemUsed_{lang}'],
    'C08': ['Code', 'Agency_{lang}', 'Type_{lang}', 'InstrumentShortName_{lang}', 'InstrumentLongName_{lang}'],
    'C11': ['CREX2', 'GRIB2_BUFR4', 'OriginatingGeneratingCentre_{lang}'],
    'C12': ['CodeFigure_OriginatingCentres', 'Name_OriginatingCentres_{lang}', 'CodeFigure_SubCentres', 'Name_SubCentres_{lang}'],
    'C13': ['CodeFigure_DataCategories', 'Name_DataCategories_{lang}', 'CodeFigure_InternationalDataSubcategories', 'Name_InternationalDataSubcategories_{lang}'],
    'C14': ['CodeFigure', 'Meaning_{lang}', 'ChemicalFormula'],
}


def load_notes(lang, base_dir):
    notes = {}
    notes_file = os.path.join(base_dir, 'notes', f'CCT_notes_{lang}.csv')
    if os.path.exists(notes_file):
        with open(notes_file, encoding='utf-8-sig') as f:
            for r in csv.DictReader(f):
                nid = r.get('noteID', '').strip()
                text = r.get(f'note_{lang}', r.get('note', '')).strip()
                if nid and text:
                    notes[nid] = text
    return notes


def set_cell(cell, text, size=SZ_DATA, bold=False):
    cell.text = ''
    p = cell.paragraphs[0]
    p.space_before = Pt(1)
    p.space_after = Pt(1)
    run = p.add_run(str(text) if text else '')
    run.font.name = FONT
    run.font.size = size
    run.font.bold = bold


def style_header(table, headers):
    for cell, hdr in zip(table.rows[0].cells, headers):
        set_cell(cell, hdr, size=SZ_HDR, bold=True)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        tcPr = cell._element.get_or_add_tcPr()
        tcPr.append(tcPr.makeelement(qn('w:shd'), {
            qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): 'D9E2F3'}))


def generate_table(lang, tbl_id, csv_path, doc, notes_db):
    with open(csv_path, encoding='utf-8-sig') as f:
        rows = list(csv.DictReader(f))
    if not rows:
        return

    # Title
    p = doc.add_paragraph()
    run = p.add_run(f'Common Code Table {tbl_id} — {TABLE_NAMES.get(tbl_id, tbl_id)}')
    run.font.name = FONT
    run.font.size = SZ_TITLE
    run.font.bold = True
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Resolve column names
    col_defs = TABLE_COLS.get(tbl_id, [])
    col_names = [c.replace('{lang}', lang) for c in col_defs]
    display_headers = [c.split('_')[0] if '_{' not in c else c.replace(f'_{lang}', '') for c in col_defs]

    table = doc.add_table(rows=1, cols=len(col_names))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    style_header(table, display_headers)

    for r in rows:
        row_cells = table.add_row().cells
        for i, col in enumerate(col_names):
            val = r.get(col, '')
            # For C06 NoteID column, resolve to full note text
            if col == 'NoteID' and val.strip() and val.strip() in notes_db:
                val = f'Note {val}: {notes_db[val.strip()]}'
            set_cell(row_cells[i], val)


def main():
    parser = argparse.ArgumentParser(description='Generate Word tables from CCT CSVs')
    parser.add_argument('lang', help='Language code (ru, fr, es, en)')
    parser.add_argument('--table', default='all', help='Table ID (C00, C14, etc.) or "all"')
    parser.add_argument('--outdir', default=None)
    args = parser.parse_args()

    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    lang_dir_map = {'en': 'english', 'fr': 'french', 'es': 'spanish', 'ru': 'russian'}
    lang_dir = os.path.join(base_dir, lang_dir_map.get(args.lang, args.lang))
    out_dir = args.outdir or os.path.join(base_dir, 'docx', args.lang)
    os.makedirs(out_dir, exist_ok=True)

    notes_db = load_notes(args.lang, base_dir)
    print(f'Loaded {len(notes_db)} notes for {args.lang}')

    tables = ['C00','C01','C02','C03','C04','C05','C06','C07','C08','C11','C12','C13','C14']
    if args.table != 'all':
        tables = [args.table]

    total = 0
    for tbl in tables:
        csv_path = os.path.join(lang_dir, f'{tbl}_{args.lang}.csv')
        if not os.path.exists(csv_path):
            continue

        doc = Document()
        section = doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Cm(29.7)
        section.page_height = Cm(21.0)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)

        generate_table(args.lang, tbl, csv_path, doc, notes_db)
        out_path = os.path.join(out_dir, f'{tbl}_{args.lang}.docx')
        doc.save(out_path)
        total += 1
        print(f'  {tbl}_{args.lang}.docx')

    print(f'\n{total} files written to {out_dir}')


if __name__ == '__main__':
    main()
