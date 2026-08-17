#!/usr/bin/env python3
"""
generate_docx.py — Build an eCH-standard .docx from a YAML content file.

Strategy: clone the eCH reference .docx (preserving all styles, headers,
footers, numbering, theme) then replace the body with generated content
using the exact eCH style names found in the reference.

Usage:
    python generate_docx.py \
        --reference  build/ech-reference.docx \
        --content    build/content.yaml \
        --output     build/index.docx \
        --images     docs/images
"""

import argparse
import copy
import os
import re
import shutil
import sys
import yaml
from lxml import etree
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ── Namespace map ──────────────────────────────────────────────────────────────
NS = {
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}


# ── Helpers ────────────────────────────────────────────────────────────────────

def clear_body(doc: Document):
    """Remove all paragraphs and tables from the document body,
    keeping the sectPr (section properties / page setup)."""
    body = doc.element.body
    sect_pr = body.find(qn('w:sectPr'))
    for child in list(body):
        if child is not sect_pr:
            body.remove(child)


def add_paragraph(doc: Document, text: str, style: str,
                  bold: bool = False, italic: bool = False,
                  align: str = None) -> None:
    """Add a paragraph with an eCH style name."""
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if align:
        alignments = {
            'right':   WD_ALIGN_PARAGRAPH.RIGHT,
            'center':  WD_ALIGN_PARAGRAPH.CENTER,
            'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
        }
        p.alignment = alignments.get(align, WD_ALIGN_PARAGRAPH.LEFT)
    return p


def add_runs(doc: Document, para, runs: list) -> None:
    """Add multiple runs (with optional bold/italic) to an existing paragraph."""
    for r in runs:
        if isinstance(r, str):
            para.add_run(r)
        else:
            run = para.add_run(r.get('text', ''))
            run.bold = r.get('bold', False)
            run.italic = r.get('italic', False)


def add_toc(doc: Document, title: str = 'Inhaltsverzeichnis') -> None:
    """Insert a TOC field that Word updates on first open."""
    # Title paragraph
    p_title = doc.add_paragraph(style='TitelInhaltsverzeichnis')
    p_title.add_run(title)

    # TOC field paragraph
    p = doc.add_paragraph()
    p.style = doc.styles['Verzeichnis1']
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), r' TOC \o "1-5" \h \z \u ')
    run = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = '[Inhaltsverzeichnis wird beim Öffnen in Word aktualisiert]'
    run.append(t)
    fld.append(run)
    p._p.append(fld)


def add_tof(doc: Document, title: str, field_name: str,
            style: str = 'Abbildungsverzeichnis') -> None:
    """Insert a TOC-style field for figures or tables."""
    p_title = doc.add_paragraph(style='Anhangberschrift')
    p_title.add_run(title)

    p = doc.add_paragraph()
    p.style = doc.styles[style]
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), rf' TOC \h \z \c "{field_name}" ')
    run = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = f'[{field_name}-Verzeichnis wird beim Öffnen in Word aktualisiert]'
    run.append(t)
    fld.append(run)
    p._p.append(fld)


def add_caption(doc: Document, caption_type: str, counter: dict,
                label: str) -> None:
    """Add an eCH-style caption with SEQ field for auto-numbering.
    caption_type: 'Abbildung' or 'Tabelle'
    """
    counter[caption_type] = counter.get(caption_type, 0) + 1
    p = doc.add_paragraph(style='Beschriftung')

    # prefix run
    r1 = p.add_run(f'{caption_type} ')

    # SEQ field
    fld = OxmlElement('w:fldSimple')
    fld.set(qn('w:instr'), f' SEQ {caption_type} \\* ARABIC ')
    r_fld = OxmlElement('w:r')
    t = OxmlElement('w:t')
    t.text = str(counter[caption_type])
    r_fld.append(t)
    fld.append(r_fld)
    p._p.append(fld)

    # suffix run
    r2 = OxmlElement('w:r')
    t2 = OxmlElement('w:t')
    t2.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t2.text = f': {label}'
    r2.append(t2)
    p._p.append(r2)


def add_metadata_table(doc: Document, metadata: dict) -> None:
    """Render the eCH cover metadata as a two-column table."""
    rows = [
        ('Name',               metadata.get('title', '')),
        ('eCH-Nummer',         metadata.get('ech_nummer', '')),
        ('Kategorie',          metadata.get('kategorie', '')),
        ('Reifegrad',          metadata.get('reifegrad', '')),
        ('Version',            metadata.get('version', '')),
        ('Status',             metadata.get('status', '')),
        ('Beschluss am',       metadata.get('beschluss', '')),
        ('Ausgabedatum',       metadata.get('ausgabe', '')),
        ('Ersetzt Version',    metadata.get('ersetzt_version', '')),
        ('Voraussetzungen',    metadata.get('voraussetzungen', '')),
        ('Beilagen',           metadata.get('beilagen', '')),
        ('Sprachen',           metadata.get('sprachen', '')),
        ('Fachgruppe',         metadata.get('fachgruppe', '')),
        ('Herausgeber / Vertrieb',
         metadata.get('herausgeber', 'Verein eCH, Affolternstrasse 52, 8050 Zürich') +
         '\n' + metadata.get('kontakt', 'T 044 388 74 64 / info@ech.ch / www.ech.ch')),
    ]

    # Column widths in twips (from reference: 2622 + 6589 = 9211)
    col_widths = [2622, 6589]
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Tabellenraster'

    # Set column widths via tblGrid
    tbl = table._tbl
    tblGrid = OxmlElement('w:tblGrid')
    for w in col_widths:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)
    # Insert tblGrid after tblPr
    tbl_pr = tbl.find(qn('w:tblPr'))
    if tbl_pr is not None:
        tbl_pr.addnext(tblGrid)

    for label, value in rows:
        row = table.add_row()
        # Label cell — Tabellentitel style
        c0 = row.cells[0]
        c0.width = Pt(col_widths[0] / 20)   # twips → points (1pt = 20 twips)
        p0 = c0.paragraphs[0]
        p0.style = doc.styles['Tabellentitel']
        p0.add_run(label)

        # Value cell — Tabellentext style
        c1 = row.cells[1]
        c1.width = Pt(col_widths[1] / 20)
        # Handle multiline values (herausgeber)
        lines = value.split('\n')
        p1 = c1.paragraphs[0]
        p1.style = doc.styles['Tabellentext']
        p1.add_run(lines[0])
        for line in lines[1:]:
            p_extra = c1.add_paragraph(style='Tabellentext')
            p_extra.add_run(line)


def add_generic_table(doc: Document, headers: list, rows: list,
                      col_widths: list = None) -> None:
    """Add a styled table with header row (Tabellentitel) and data rows (Tabellentext)."""
    n_cols = len(headers)
    table = doc.add_table(rows=0, cols=n_cols)
    table.style = 'Tabellenraster'

    if col_widths:
        tbl = table._tbl
        tblGrid = OxmlElement('w:tblGrid')
        for w in col_widths:
            gc = OxmlElement('w:gridCol')
            gc.set(qn('w:w'), str(w))
            tblGrid.append(gc)
        tbl_pr = tbl.find(qn('w:tblPr'))
        if tbl_pr is not None:
            tbl_pr.addnext(tblGrid)

    # Header row (only if headers are provided)
    if headers:
        hrow = table.add_row()
        for i, h in enumerate(headers):
            c = hrow.cells[i]
            p = c.paragraphs[0]
            p.style = doc.styles['Tabellentitel']
            p.add_run(h)

    # Data rows
    for data_row in rows:
        drow = table.add_row()
        for i, val in enumerate(data_row):
            c = drow.cells[i]
            p = c.paragraphs[0]
            p.style = doc.styles['Tabellentext']
            p.add_run(str(val))


def add_list_item(doc: Document, text: str, style_name: str) -> None:
    """Add a list item using eCH list styles."""
    p = doc.add_paragraph(style=style_name)
    p.add_run(text)


def add_image(doc: Document, image_path: str, width_cm: float = 6.0,
              align: str = 'right') -> None:
    """Add an image paragraph, right-aligned per eCH template."""
    if not os.path.exists(image_path):
        doc.add_paragraph(f'[Bild nicht gefunden: {image_path}]', style='Standard')
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT if align == 'right' else WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run()
    run.add_picture(image_path, width=Cm(width_cm))


def page_break(doc: Document) -> None:
    p = doc.add_paragraph()
    run = p.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._r.append(br)


# ── Content renderer ───────────────────────────────────────────────────────────

def render_blocks(doc: Document, blocks: list, images_dir: str,
                  caption_counter: dict) -> None:
    """Recursively render a list of content blocks."""
    for block in blocks:
        btype = block.get('type', 'paragraph')

        if btype == 'paragraph':
            p = doc.add_paragraph(style=block.get('style', 'Standard'))
            text = block.get('text', '')
            runs = block.get('runs', None)
            if runs:
                add_runs(doc, p, runs)
            else:
                r = p.add_run(text)
                r.bold = block.get('bold', False)
                r.italic = block.get('italic', False)

        elif btype == 'heading':
            level = block.get('level', 1)
            styles = {1: 'berschrift1', 2: 'berschrift2', 3: 'berschrift3',
                      4: 'berschrift4', 5: 'berschrift5'}
            style = styles.get(level, 'berschrift1')
            p = doc.add_paragraph(style=style)
            p.add_run(block.get('text', ''))

        elif btype == 'anhang_heading':
            p = doc.add_paragraph(style='Anhangberschrift')
            p.add_run(block.get('text', ''))

        elif btype == 'nebentitel':
            p = doc.add_paragraph(style='Nebentitel')
            p.add_run(block.get('text', ''))

        elif btype == 'subtitle':
            p = doc.add_paragraph(style='Untertitel')
            p.add_run(block.get('text', ''))

        elif btype == 'austauschformat':
            p = doc.add_paragraph(style='Austauschformat')
            p.add_run(block.get('text', ''))

        elif btype == 'list_item':
            style_map = {
                'bullet1': 'Aufzhlung',
                'bullet2': 'Liste-',
                'number1': 'Liste1',
                'alpha1':  'Listea',
            }
            style = style_map.get(block.get('list_style', 'bullet1'), 'Aufzhlung')
            p = doc.add_paragraph(style=style)
            p.add_run(block.get('text', ''))

        elif btype == 'image':
            add_image(doc, os.path.join(images_dir, block.get('file', '')),
                      width_cm=block.get('width_cm', 6.0),
                      align=block.get('align', 'right'))
            if 'caption' in block:
                add_caption(doc, 'Abbildung', caption_counter, block['caption'])

        elif btype == 'table':
            headers = block.get('headers', [])
            rows = block.get('rows', [])
            col_widths = block.get('col_widths', None)
            add_generic_table(doc, headers, rows, col_widths)
            if 'caption' in block:
                add_caption(doc, 'Tabelle', caption_counter, block['caption'])

        elif btype == 'toc':
            add_toc(doc, block.get('title', 'Inhaltsverzeichnis'))

        elif btype == 'page_break':
            page_break(doc)

        elif btype == 'empty':
            doc.add_paragraph(style='Standard')

        elif btype == 'section':
            render_blocks(doc, block.get('blocks', []), images_dir, caption_counter)


# ── Cover page ────────────────────────────────────────────────────────────────

def render_cover(doc: Document, metadata: dict) -> None:
    """Render the cover page: title + metadata table + Zusammenfassung + Hinweis."""
    # Document title (Titel style — no heading numbering)
    p_title = doc.add_paragraph(style='Titel')
    p_title.add_run(f"{metadata.get('ech_nummer', '')} – {metadata.get('title', '')}")

    # Metadata table
    add_metadata_table(doc, metadata)

    # Empty line
    doc.add_paragraph(style='Standard')

    # Zusammenfassung (Nebentitel — not a numbered heading)
    p_zus = doc.add_paragraph(style='Nebentitel')
    p_zus.add_run('Zusammenfassung')
    p_body = doc.add_paragraph(style='Standard')
    p_body.add_run(metadata.get('zusammenfassung',
                                '<Kurze Zusammenfassung des Zwecks des Dokuments>'))
    p_body.runs[0].italic = True

    # Empty Nebentitel spacer (matches template para [52])
    doc.add_paragraph(style='Nebentitel')

    # Hinweis (Nebentitel — not numbered)
    p_hw = doc.add_paragraph(style='Nebentitel')
    p_hw.add_run('Hinweis')
    p_hw_body = doc.add_paragraph(style='Standard')
    p_hw_body.add_run(
        'Im vorliegenden Dokument wird bei der Bezeichnung von Personen eine '
        'geschlechtsneutrale Formulierung verwendet. Basis bildet der Leitfaden '
        'der Bundeskanzlei. Je nach Situation kommen Paarformen (Bürgerinnen und '
        'Bürger), geschlechtsabstrakte Formen (versicherte Person), '
        'geschlechtsneutrale Formen (Versicherte) oder Umschreibungen ohne '
        'Personenbezug zum Einsatz. Das generische Maskulin (Bürger) ist nicht '
        'zulässig. Vollformen werden in fortlaufenden Texten verwendet. In '
        'verknappten Textpassagen, namentlich in Tabellen, können Kurzformen '
        'verwendet werden (Referent/in). Genderstern und ähnliche Schreibweisen '
        'werden nicht verwendet.'
    )


# ── Appendices ────────────────────────────────────────────────────────────────

def render_appendices(doc: Document, appendices: list,
                      images_dir: str, caption_counter: dict) -> None:
    labels = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
    for i, app in enumerate(appendices):
        label = labels[i] if i < len(labels) else str(i + 1)
        title = app.get('title', f'Anhang {label}')
        heading = f'Anhang {label} – {title}'

        app_type = app.get('type', 'generic')

        if app_type == 'abbildungsverzeichnis':
            add_tof(doc, heading, 'Abbildung', style='Abbildungsverzeichnis')
        elif app_type == 'tabellenverzeichnis':
            add_tof(doc, heading, 'Tabelle', style='Abbildungsverzeichnis')
        else:
            p = doc.add_paragraph(style='Anhangberschrift')
            p.add_run(heading)
            render_blocks(doc, app.get('blocks', []), images_dir, caption_counter)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Generate eCH-standard .docx')
    parser.add_argument('--reference', required=True, help='Path to eCH reference .docx')
    parser.add_argument('--content',   required=True, help='Path to content YAML file')
    parser.add_argument('--output',    required=True, help='Output .docx path')
    parser.add_argument('--images',    default='docs/images', help='Images directory')
    args = parser.parse_args()

    # Load content
    with open(args.content, encoding='utf-8') as f:
        content = yaml.safe_load(f)

    # Clone reference to preserve all styles/headers/footers
    shutil.copy2(args.reference, args.output)
    doc = Document(args.output)

    # Clear body
    clear_body(doc)

    metadata = content.get('metadata', {})
    caption_counter = {}

    # 1. Cover page
    render_cover(doc, metadata)

    # 2. TOC
    add_toc(doc)

    # 3. Body chapters
    render_blocks(doc, content.get('chapters', []),
                  args.images, caption_counter)

    # 4. Appendices
    render_appendices(doc, content.get('appendices', []),
                      args.images, caption_counter)

    doc.save(args.output)
    print(f'Generated: {args.output}')


if __name__ == '__main__':
    main()
