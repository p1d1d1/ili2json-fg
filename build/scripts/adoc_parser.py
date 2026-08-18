"""
adoc_parser.py — Minimal AsciiDoc parser for eCH document generation.

Supported constructs:
  Headings         == through ======
  Paragraphs       plain text (with inline bold/italic)
  Bullet lists     * item / ** item / *** item  (space after * required)
  Ordered lists    . item / .. item (numeric); [loweralpha] . item (alpha)
  Images           image::file[opts]
  Tables           |=== blocks (with optional .Caption and [cols=...])
  Roles            [.austauschformat], [.untertitel] before a paragraph
  Inline markup    *bold*, _italic_, *_bold-italic_*
  Skipped          // comments, :attr: lines, include::, [[anchors]], [appendix]

parse_adoc() accepts an optional skip_first_heading=True flag used when
the file is loaded as an appendix body (render_appendices provides the heading).
"""

import re
import os


# ── Inline markup ──────────────────────────────────────────────────────────────

def parse_inline(text):
    """
    Parse inline bold/italic and return a str (if plain) or list of run dicts.
    """
    pattern = re.compile(
        r'(\*_(?P<bi1>.*?)_\*)'
        r'|(_\*(?P<bi2>.*?)\*_)'
        r'|(\*(?P<b>[^*\s][^*]*?)\*)'   # *bold* — requires non-space after opening *
        r'|(_(?P<i>[^_\s][^_]*?)_)',     # _italic_
        re.DOTALL
    )
    runs = []
    last = 0
    for m in pattern.finditer(text):
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
    if len(runs) == 1 and not runs[0]['bold'] and not runs[0]['italic']:
        return runs[0]['text']
    return runs


def runs_or_text(text):
    parsed = parse_inline(text)
    if isinstance(parsed, str):
        return {'text': parsed}
    return {'runs': parsed}


# ── Line classification helpers ────────────────────────────────────────────────

def is_heading(line):
    """Return (level, title) or None."""
    m = re.match(r'^(={2,6})\s+(.+)$', line.strip())
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None

def is_bullet(line):
    """Return (depth, text) for '* text' / '** text' lines, else None.
    Requires a space after the asterisk(s) to distinguish from inline bold.
    """
    m = re.match(r'^(\*{1,3})\s+(.+)$', line.strip())
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None

def is_ordered(line):
    """Return (depth, text) for '. text' / '.. text' lines, else None."""
    m = re.match(r'^(\.{1,2})\s+(.+)$', line.strip())
    if m:
        return len(m.group(1)), m.group(2).strip()
    return None

def is_block_boundary(line):
    """True if this line should stop paragraph continuation."""
    s = line.strip()
    if not s:
        return True
    if s.startswith('//'):
        return True
    if re.match(r'^={2,6}\s', s):          # heading
        return True
    if re.match(r'^\*{1,3}\s', s):         # bullet list (space required)
        return True
    if re.match(r'^\.{1,2}\s', s):         # ordered list
        return True
    if s == '|===':                         # table
        return True
    if re.match(r'^\[.*\]$', s):           # block attribute
        return True
    if s.startswith('image::'):             # image
        return True
    if s.startswith('include::'):           # include
        return True
    if re.match(r'^\[\[.*\]\]$', s):       # anchor
        return True
    return False


# ── Table parser ───────────────────────────────────────────────────────────────

def parse_table(lines, start_idx, caption=None, col_widths=None, has_header=False):
    rows = []
    i = start_idx + 1
    while i < len(lines) and lines[i].strip() != '|===':
        line = lines[i].strip()
        if line and line.startswith('|'):
            cells = [c.strip() for c in line[1:].split('|')]
            rows.append(cells)
        i += 1

    block = {'type': 'table', 'headers': [], 'rows': rows}
    if has_header and rows:
        block['headers'] = rows[0]
        block['rows'] = rows[1:]
    if caption:
        block['caption'] = re.sub(r'^(Tabelle|Table)\s*:\s*', '', caption)
    if col_widths:
        block['col_widths'] = col_widths
    return block, i   # i points to closing |===


# ── Main parser ────────────────────────────────────────────────────────────────

def parse_adoc(filepath, images_dir=None, skip_first_heading=False):
    """
    Parse an AsciiDoc file and return a list of block dicts.

    skip_first_heading=True: skip the first == heading encountered.
    Used when the file is loaded as appendix body content — the heading
    is provided by render_appendices() and must not be duplicated.
    """
    with open(filepath, encoding='utf-8') as f:
        lines = [l.rstrip('\n') for l in f]

    blocks = []
    i = 0
    pending_role        = None
    pending_caption     = None
    pending_col_widths  = None
    pending_has_header  = False
    pending_alpha       = False
    first_heading_seen  = False

    while i < len(lines):
        line     = lines[i]
        stripped = line.strip()

        # ── Blank line ────────────────────────────────────────
        if not stripped:
            pending_role = None
            i += 1
            continue

        # ── Comment ───────────────────────────────────────────
        if stripped.startswith('//'):
            i += 1
            continue

        # ── Anchor [[...]] ────────────────────────────────────
        if re.match(r'^\[\[.*\]\]$', stripped):
            i += 1
            continue

        # ── Attribute definition :key: val ───────────────────
        if re.match(r'^:[^:]+:', stripped):
            i += 1
            continue

        # ── include:: ─────────────────────────────────────────
        if stripped.startswith('include::'):
            i += 1
            continue

        # ── Block attribute line [something] ──────────────────
        if re.match(r'^\[([^\]]*)\]$', stripped):
            attr = re.match(r'^\[([^\]]*)\]$', stripped).group(1)
            if attr == 'appendix':
                # Skip [appendix] role — heading handled externally
                i += 1
                continue
            elif attr.startswith('.'):
                pending_role = attr[1:]
            elif 'loweralpha' in attr:
                pending_alpha = True
            elif 'header' in attr:
                pending_has_header = True
            # Parse col widths from [cols="..."] 
            cols_m = re.search(r'cols=["\']([^"\']+)["\']', attr)
            if cols_m:
                parts = [p.strip() for p in cols_m.group(1).split(',')]
                widths = []
                for p in parts:
                    m2 = re.match(r'^(\d+)', p)
                    if m2:
                        widths.append(int(m2.group(1)) * 500)
                    else:
                        widths.append(3000)   # default for ~,^,<,> etc.
                if widths:
                    pending_col_widths = widths
            i += 1
            continue

        # ── Caption (.text) ───────────────────────────────────
        if stripped.startswith('.') and not re.match(r'^\.\.\s', stripped):
            pending_caption = stripped[1:].strip()
            i += 1
            continue

        # ── Heading ───────────────────────────────────────────
        h = is_heading(stripped)
        if h:
            level, title = h
            if skip_first_heading and not first_heading_seen:
                first_heading_seen = True
                i += 1
                continue
            first_heading_seen = True
            # Clamp level: adoc == is h1 in chapter context
            blocks.append({'type': 'heading', 'level': level - 1, 'text': title})
            pending_role = None
            pending_caption = None
            i += 1
            continue

        # ── Image ─────────────────────────────────────────────
        m_img = re.match(r'^image::([^\[]+)\[([^\]]*)\]$', stripped)
        if m_img:
            img_file = m_img.group(1).strip()
            img_opts = m_img.group(2)
            caption  = pending_caption
            if not caption:
                cm = re.search(r'caption=["\']([^"\']+)["\']', img_opts)
                if cm:
                    caption = cm.group(1)
            wm = re.search(r'width=(\d+)', img_opts)
            width_cm = float(wm.group(1)) / 37.8 if wm else 6.0
            width_cm = min(max(width_cm, 3.0), 14.0)
            block = {'type': 'image', 'file': img_file,
                     'width_cm': width_cm, 'align': 'right'}
            if caption:
                block['caption'] = re.sub(r'^(Abbildung|Figure)\s*:\s*', '', caption)
            blocks.append(block)
            pending_caption = None
            pending_role    = None
            i += 1
            continue

        # ── Table ─────────────────────────────────────────────
        if stripped == '|===':
            tbl, end_i = parse_table(lines, i,
                                     caption=pending_caption,
                                     col_widths=pending_col_widths,
                                     has_header=pending_has_header)
            blocks.append(tbl)
            pending_caption    = None
            pending_col_widths = None
            pending_has_header = False
            i = end_i + 1
            continue

        # ── Bullet list ───────────────────────────────────────
        b_item = is_bullet(stripped)
        if b_item:
            depth, text = b_item
            style_map = {1: 'bullet1', 2: 'bullet2', 3: 'bullet3'}
            block = {'type': 'list_item', 'list_style': style_map[depth]}
            block.update(runs_or_text(text))
            blocks.append(block)
            pending_role = None
            i += 1
            continue

        # ── Ordered list ──────────────────────────────────────
        o_item = is_ordered(stripped)
        if o_item:
            depth, text = o_item
            style = 'alpha1' if (pending_alpha and depth == 1) else 'number1'
            if depth == 1:
                pending_alpha = False
            block = {'type': 'list_item', 'list_style': style}
            block.update(runs_or_text(text))
            blocks.append(block)
            pending_role = None
            i += 1
            continue

        # ── Admonition block (NOTE / TIP / WARNING …) ─────────
        if stripped in ('NOTE', 'TIP', 'WARNING', 'IMPORTANT', 'CAUTION') \
                and i + 1 < len(lines) and lines[i + 1].strip() == '====':
            i += 2
            admon = []
            while i < len(lines) and lines[i].strip() != '====':
                admon.append(lines[i].strip())
                i += 1
            text = ' '.join(l for l in admon if l)
            block = {'type': 'paragraph'}
            block.update(runs_or_text(text))
            blocks.append(block)
            i += 1
            continue

        # ── Plain paragraph ───────────────────────────────────
        # Collect continuation lines until a block boundary
        para_lines = [stripped]
        i += 1
        while i < len(lines) and not is_block_boundary(lines[i]):
            para_lines.append(lines[i].strip())
            i += 1

        text = ' '.join(para_lines)

        if pending_role == 'austauschformat':
            blocks.append({'type': 'austauschformat', 'text': text})
        elif pending_role == 'untertitel':
            blocks.append({'type': 'subtitle', 'text': text})
        else:
            block = {'type': 'paragraph'}
            block.update(runs_or_text(text))
            blocks.append(block)

        pending_role    = None
        pending_caption = None

    return blocks
