"""
adoc_parser.py — Minimal AsciiDoc parser for eCH document generation.

Converts a subset of AsciiDoc syntax into the block list that
generate_docx.py understands. Supported constructs:

  Headings         == through ======
  Paragraphs       plain text (with inline bold/italic)
  Bullet lists     * / ** / ***
  Ordered lists    . / .. (numeric)  [loweralpha] . (alpha)
  Images           image::file[alt,caption="..."]
  Tables           |=== blocks (with optional .Caption and [cols=...])
  Roles            [.austauschformat], [.untertitel] before a paragraph
  Inline markup    *bold*, _italic_, *_bold-italic_*
  Comments         // lines — skipped
  Attribute lines  :key: value — skipped
  Include lines    include:: — skipped
"""

import re
import os


# ── Inline markup ──────────────────────────────────────────────────────────────

def parse_inline(text):
    """
    Parse inline bold/italic markup and return a list of run dicts:
      [{'text': str, 'bold': bool, 'italic': bool}, ...]
    Handles: *bold*, _italic_, *_both_*, _*both*_
    """
    # Tokenise by bold+italic, bold, italic markers
    pattern = re.compile(
        r'(\*_(?P<bi1>.*?)_\*)'     # *_bold-italic_*
        r'|(_\*(?P<bi2>.*?)\*_)'    # _*bold-italic*_
        r'|(\*(?P<b>.*?)\*)'        # *bold*
        r'|(_(?P<i>[^_]+)_)',       # _italic_
        re.DOTALL
    )
    runs = []
    last = 0
    for m in pattern.finditer(text):
        # plain text before this match
        if m.start() > last:
            runs.append({'text': text[last:m.start()], 'bold': False, 'italic': False})
        bi1 = m.group('bi1')
        bi2 = m.group('bi2')
        b   = m.group('b')
        i   = m.group('i')
        if bi1 is not None:
            runs.append({'text': bi1, 'bold': True, 'italic': True})
        elif bi2 is not None:
            runs.append({'text': bi2, 'bold': True, 'italic': True})
        elif b is not None:
            runs.append({'text': b, 'bold': True, 'italic': False})
        elif i is not None:
            runs.append({'text': i, 'bold': False, 'italic': True})
        last = m.end()
    if last < len(text):
        runs.append({'text': text[last:], 'bold': False, 'italic': False})
    # Simplify: if single plain run, just use text string
    if len(runs) == 1 and not runs[0]['bold'] and not runs[0]['italic']:
        return runs[0]['text']
    return runs


def runs_or_text(text):
    """Return {'text': ...} or {'runs': [...]} depending on inline content."""
    parsed = parse_inline(text)
    if isinstance(parsed, str):
        return {'text': parsed}
    return {'runs': parsed}


# ── Table parser ───────────────────────────────────────────────────────────────

def parse_table(lines, start_idx, caption=None, col_widths=None):
    """Parse a |=== table block, return a table block dict."""
    rows = []
    headers = []
    has_header = False
    i = start_idx + 1  # skip opening |===

    while i < len(lines) and lines[i].strip() != '|===':
        line = lines[i].strip()
        if not line:
            i += 1
            continue
        if line.startswith('|'):
            # Split cells by | (skip leading |)
            cells = [c.strip() for c in line[1:].split('|')]
            rows.append(cells)
        i += 1

    # If [options="header"] was set, first row is header
    block = {
        'type': 'table',
        'headers': [],
        'rows': rows,
    }
    if has_header and rows:
        block['headers'] = rows[0]
        block['rows'] = rows[1:]
    if caption:
        block['caption'] = caption
    if col_widths:
        block['col_widths'] = col_widths

    return block, i  # i points to closing |===


# ── Heading level helpers ──────────────────────────────────────────────────────

def heading_level(line):
    """Return (level, title) for a heading line, or None."""
    m = re.match(r'^(={2,6})\s+(.+)$', line)
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None


# ── Main parser ────────────────────────────────────────────────────────────────

def parse_adoc(filepath, images_dir=None):
    """
    Parse an AsciiDoc file and return a list of block dicts
    suitable for generate_docx.py's render_blocks().
    """
    with open(filepath, encoding='utf-8') as f:
        raw_lines = f.readlines()

    lines = [l.rstrip('\n') for l in raw_lines]
    blocks = []
    i = 0
    pending_role = None       # [.role] annotation pending for next block
    pending_caption = None    # .Caption pending for next block
    pending_col_widths = None # [cols=...] pending for next table
    pending_ordered_alpha = False
    caption_counter = {'Abbildung': 0, 'Tabelle': 0}

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        # ── Skip blank lines ──────────────────────────────────
        if not stripped:
            pending_role = None
            i += 1
            continue

        # ── Skip comments ─────────────────────────────────────
        if stripped.startswith('//'):
            i += 1
            continue

        # ── Skip attribute definitions (:key: val) ────────────
        if re.match(r'^:[^:]+:', stripped):
            i += 1
            continue

        # ── Skip include directives ───────────────────────────
        if stripped.startswith('include::'):
            i += 1
            continue

        # ── Block attribute line [.role] or [options] etc. ────
        m_attr = re.match(r'^\[([^\]]+)\]$', stripped)
        if m_attr:
            attr = m_attr.group(1)
            if attr.startswith('.'):
                pending_role = attr[1:]  # e.g. 'austauschformat'
            elif 'loweralpha' in attr:
                pending_ordered_alpha = True
            elif attr.startswith('cols=') or 'cols=' in attr:
                # extract col widths if numeric
                col_m = re.findall(r'(\d+)', attr)
                if col_m:
                    pending_col_widths = [int(c) * 500 for c in col_m]
            # options="header" handled at table parse time
            i += 1
            continue

        # ── Caption line (.Caption text) ──────────────────────
        if stripped.startswith('.') and not stripped.startswith('..'):
            pending_caption = stripped[1:].strip()
            i += 1
            continue

        # ── Heading ───────────────────────────────────────────
        h = heading_level(stripped)
        if h:
            level, title = h
            blocks.append({'type': 'heading', 'level': level - 1, 'text': title})
            # level-1 because adoc uses == for h1 but our schema uses level 1
            pending_role = None
            pending_caption = None
            i += 1
            continue

        # ── Image ─────────────────────────────────────────────
        m_img = re.match(r'^image::([^\[]+)\[([^\]]*)\]$', stripped)
        if m_img:
            img_file = m_img.group(1).strip()
            img_opts = m_img.group(2)
            caption = pending_caption
            if not caption:
                # Try caption= inside the brackets
                cm = re.search(r'caption="([^"]+)"', img_opts)
                if cm:
                    caption = cm.group(1)
            # Width from options e.g. width=200
            wm = re.search(r'width=(\d+)', img_opts)
            width_cm = float(wm.group(1)) / 37.8 if wm else 6.0  # px→cm rough
            width_cm = min(max(width_cm, 3.0), 14.0)
            block = {
                'type': 'image',
                'file': img_file,
                'width_cm': width_cm,
                'align': 'right',
            }
            if caption:
                block['caption'] = re.sub(r'^(Abbildung|Figure)\s*:\s*', '', caption)
            blocks.append(block)
            pending_caption = None
            pending_role = None
            i += 1
            continue

        # ── Table ─────────────────────────────────────────────
        if stripped == '|===':
            tbl_block, end_i = parse_table(lines, i,
                                           caption=pending_caption,
                                           col_widths=pending_col_widths)
            # Detect header row from [options="header"] already consumed
            # Check if the line before |=== contained options="header"
            # Simple heuristic: if pending_col_widths had options, first row = header
            # Better: look back one line
            if end_i > i:
                # check for options="header" in recent attr lines
                for back in range(max(0, i-3), i):
                    if 'options="header"' in lines[back] or "options='header'" in lines[back]:
                        if tbl_block['rows']:
                            tbl_block['headers'] = tbl_block['rows'][0]
                            tbl_block['rows'] = tbl_block['rows'][1:]
                        break
            # Clean up caption prefix from .Caption line
            if tbl_block.get('caption'):
                tbl_block['caption'] = re.sub(
                    r'^(Tabelle|Table)\s*:\s*', '', tbl_block['caption'])
            blocks.append(tbl_block)
            pending_caption = None
            pending_col_widths = None
            i = end_i + 1
            continue

        # ── Bullet list item * / ** / *** ─────────────────────
        m_bullet = re.match(r'^(\*{1,3})\s+(.+)$', stripped)
        if m_bullet:
            depth = len(m_bullet.group(1))
            text  = m_bullet.group(2).strip()
            style_map = {1: 'bullet1', 2: 'bullet2', 3: 'bullet3'}
            b = {'type': 'list_item', 'list_style': style_map[depth]}
            b.update(runs_or_text(text))
            blocks.append(b)
            pending_role = None
            i += 1
            continue

        # ── Ordered list item . / .. ──────────────────────────
        m_ordered = re.match(r'^(\.{1,2})\s+(.+)$', stripped)
        if m_ordered:
            depth = len(m_ordered.group(1))
            text  = m_ordered.group(2).strip()
            if pending_ordered_alpha and depth == 1:
                style = 'alpha1'
                pending_ordered_alpha = False
            else:
                style = 'number1'
            b = {'type': 'list_item', 'list_style': style}
            b.update(runs_or_text(text))
            blocks.append(b)
            pending_role = None
            i += 1
            continue

        # ── NOTE / TIP / WARNING admonition blocks ────────────
        if stripped in ('NOTE', 'TIP', 'WARNING', 'IMPORTANT', 'CAUTION'):
            # next line should be ====
            if i + 1 < len(lines) and lines[i+1].strip() == '====':
                i += 2  # skip label and ====
                admon_lines = []
                while i < len(lines) and lines[i].strip() != '====':
                    admon_lines.append(lines[i].strip())
                    i += 1
                text = ' '.join(l for l in admon_lines if l)
                b = {'type': 'paragraph', 'style': 'Hinweis'}
                b.update(runs_or_text(text))
                blocks.append(b)
                i += 1  # skip closing ====
                continue

        # ── Plain paragraph (may have pending role) ───────────
        # Collect continuation lines
        para_lines = [stripped]
        i += 1
        while i < len(lines) and lines[i].strip() and \
              not lines[i].strip().startswith('=') and \
              not lines[i].strip().startswith('*') and \
              not lines[i].strip().startswith('.') and \
              not lines[i].strip().startswith('|') and \
              not lines[i].strip().startswith('[') and \
              not lines[i].strip().startswith('image::') and \
              not lines[i].strip().startswith('//'):
            para_lines.append(lines[i].strip())
            i += 1

        text = ' '.join(para_lines)

        if pending_role == 'austauschformat':
            blocks.append({'type': 'austauschformat', 'text': text})
        elif pending_role == 'untertitel':
            blocks.append({'type': 'subtitle', 'text': text})
        else:
            b = {'type': 'paragraph'}
            b.update(runs_or_text(text))
            blocks.append(b)

        pending_role = None
        pending_caption = None

    return blocks
