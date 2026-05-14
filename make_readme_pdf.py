#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate README.pdf from README.md  –  Windows Event Analyzer"""

import re
import os

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib.colors import HexColor, black, white
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, Preformatted, KeepTogether,
)
from reportlab.lib.enums import TA_LEFT, TA_CENTER

PAGE_W, PAGE_H = A4
MARGIN      = 2.0 * cm
DOC_WIDTH   = PAGE_W - 2 * MARGIN

# ── Colours ─────────────────────────────────────────────────────────────────
C_BLUE   = HexColor('#1e3a5f')
C_BLUE2  = HexColor('#2d6a9f')
C_LTBLUE = HexColor('#dce8f5')
C_GRAY   = HexColor('#888888')
C_BGCODE = HexColor('#f4f4f4')
C_BGTBL  = HexColor('#eaf1fb')

# ── Emoji → ASCII map ────────────────────────────────────────────────────────
EMOJI_MAP = {
    '⚠️': '[!]',      # ⚠️
    '⚠':       '[!]',      # ⚠
    '✅':       '[OK]',     # ✅
    '\U0001f534':   '[ROT]',    # 🔴
    '\U0001f497':   '[LACHS]',  # 🩷
    '\U0001f7e1':   '[GELB]',   # 🟡
    '\U0001f4cb':   '',         # 📋
    '\U0001f4c2':   '',         # 📂
    '\U0001f4be':   '',         # 💾
    '⟳':       '',         # ⟳
    '\U0001f50e':   '',         # 🔎
    '\U0001f9ea':   '',         # 🧪
    '\U0001f4c8':   '',         # 📈
    '⚙':       '',         # ⚙
    '⚙️': '',         # ⚙️
    '\U0001f50d':   '',         # 🔍
    '\U0001f4ca':   '',         # 📊
    '➕':       '+',        # ➕
    '\U0001f5a5️': '',     # 🖥️
    '\U0001f5a5':   '',         # 🖥
    '\U0001f50c':   '',         # 🔌
    '★':       '*',        # ★
    '◈':       'o',        # ◈
    '\U0001f504':   '',         # 🔄
    '✔️': '[x]',      # ✔️
    '✔':       '[x]',      # ✔
    '✖️': '[ ]',      # ✖️
    '✖':       '[ ]',      # ✖
    '☑️': '[x]',      # ☑️
    '☑':       '[x]',      # ☑
    '☐':       '[ ]',      # ☐
    '●':       'o',        # ●
    '▲':       '^',        # ▲
    '○':       'o',        # ○
}

def strip_emojis(text):
    for emoji, replacement in EMOJI_MAP.items():
        text = text.replace(emoji, replacement)
    # Remove remaining surrogate / high-codepoint chars Helvetica can't render
    text = re.sub(r'[\U00010000-\U0010ffff]', '', text)
    return text


# ── Styles ───────────────────────────────────────────────────────────────────
def build_styles():
    return {
        'Title': ParagraphStyle('MyTitle',
            fontSize=20, leading=26, textColor=C_BLUE,
            spaceAfter=6, fontName='Helvetica-Bold'),

        'Badge': ParagraphStyle('Badge',
            fontSize=8, leading=12, textColor=C_BLUE2,
            spaceAfter=8, fontName='Helvetica'),

        'Desc': ParagraphStyle('Desc',
            fontSize=10, leading=14, textColor=HexColor('#333333'),
            spaceAfter=10, fontName='Helvetica-Oblique'),

        'H2': ParagraphStyle('H2',
            fontSize=13, leading=18, textColor=C_BLUE,
            spaceBefore=10, spaceAfter=4, fontName='Helvetica-Bold'),

        'H3': ParagraphStyle('H3',
            fontSize=10, leading=14, textColor=C_BLUE2,
            spaceBefore=8, spaceAfter=3, fontName='Helvetica-Bold'),

        'Normal': ParagraphStyle('MyNormal',
            fontSize=9, leading=13, textColor=black,
            spaceAfter=4, fontName='Helvetica'),

        'Bullet': ParagraphStyle('Bullet',
            fontSize=9, leading=13, textColor=black,
            leftIndent=12, spaceAfter=2, fontName='Helvetica'),

        'Bullet2': ParagraphStyle('Bullet2',
            fontSize=9, leading=13, textColor=black,
            leftIndent=24, spaceAfter=2, fontName='Helvetica'),

        'BlockQuote': ParagraphStyle('BlockQuote',
            fontSize=9, leading=13, textColor=HexColor('#444444'),
            leftIndent=14, spaceAfter=4, fontName='Helvetica-Oblique'),

        'Code': ParagraphStyle('Code',
            fontSize=7.5, leading=10.5, textColor=black,
            fontName='Courier', backColor=C_BGCODE,
            leftIndent=6, rightIndent=6,
            spaceAfter=6, spaceBefore=4,
            borderPadding=(4, 4, 4, 4)),

        'TH': ParagraphStyle('TH',
            fontSize=8, leading=11, textColor=white,
            fontName='Helvetica-Bold'),

        'TD': ParagraphStyle('TD',
            fontSize=8, leading=11, textColor=black,
            fontName='Helvetica'),

        'Footer': ParagraphStyle('Footer',
            fontSize=8, leading=11, textColor=C_GRAY,
            alignment=TA_CENTER, fontName='Helvetica-Oblique'),
    }


# ── Inline markdown → ReportLab XML ─────────────────────────────────────────
def xml_escape(text):
    text = text.replace('&', '&amp;')
    text = text.replace('<', '&lt;')
    text = text.replace('>', '&gt;')
    return text

def inline(text):
    """Escape XML then convert inline markdown to RL tags."""
    text = xml_escape(text)
    text = re.sub(r'\*\*(.+?)\*\*', r'<b>\1</b>', text)
    text = re.sub(r'__(.+?)__',     r'<b>\1</b>', text)
    text = re.sub(r'\*(.+?)\*',     r'<i>\1</i>', text)
    text = re.sub(r'`([^`]+)`',
                  r'<font name="Courier" size="7.5">\1</font>', text)
    return text


# ── Markdown table → RL Table ────────────────────────────────────────────────
def parse_md_table(table_lines):
    """Return list-of-lists of cell strings (header in row 0)."""
    rows = []
    for line in table_lines:
        stripped = line.strip()
        if not stripped.startswith('|'):
            continue
        if re.match(r'^\|[\s\-:|]+\|', stripped):   # separator row
            continue
        cells = [c.strip() for c in stripped.strip('|').split('|')]
        rows.append(cells)
    return rows

def col_widths_for(ncols, rows):
    """Heuristic column widths depending on column count."""
    if ncols == 2:
        return [DOC_WIDTH * 0.33, DOC_WIDTH * 0.67]
    if ncols == 3:
        # Detect if last column is short (e.g. "✅ Sicher")
        last_max = max(len(r[2]) if len(r) > 2 else 0 for r in rows)
        if last_max < 20:
            return [DOC_WIDTH * 0.26, DOC_WIDTH * 0.57, DOC_WIDTH * 0.17]
        return [DOC_WIDTH / 3] * 3
    return [DOC_WIDTH / ncols] * ncols

def build_table(rows, styles):
    if not rows:
        return None

    ncols  = max(len(r) for r in rows)
    cws    = col_widths_for(ncols, rows)

    # Pad rows to uniform width
    padded = [r + [''] * (ncols - len(r)) for r in rows]

    rl_rows = []
    for i, row in enumerate(padded):
        st = styles['TH'] if i == 0 else styles['TD']
        rl_rows.append([Paragraph(inline(cell), st) for cell in row])

    t = Table(rl_rows, colWidths=cws, repeatRows=1,
              hAlign='LEFT', spaceBefore=4, spaceAfter=6)
    t.setStyle(TableStyle([
        ('BACKGROUND',    (0, 0), (-1,  0), C_BLUE),
        ('TEXTCOLOR',     (0, 0), (-1,  0), white),
        ('ROWBACKGROUNDS',(0, 1), (-1, -1), [white, C_BGTBL]),
        ('GRID',          (0, 0), (-1, -1), 0.35, HexColor('#cccccc')),
        ('VALIGN',        (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING',    (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LEFTPADDING',   (0, 0), (-1, -1), 5),
        ('RIGHTPADDING',  (0, 0), (-1, -1), 5),
    ]))
    return t


# ── Main parser ──────────────────────────────────────────────────────────────
def md_to_flowables(md_text, S):
    """Convert full markdown string to list of RL flowables."""
    # Pre-process
    md_text = strip_emojis(md_text)

    lines = md_text.split('\n')
    result = []
    i      = 0
    N      = len(lines)

    while i < N:
        line = lines[i]

        # ── Badge line (shields.io images) ──────────────────────────────────
        if re.match(r'^!\[.*\]\(https://img\.shields\.io', line):
            badges = re.findall(
                r'!\[([^\]]+)\]\(https://img\.shields\.io/badge/([^)]+)\)', line)
            parts = []
            for alt, path in badges:
                segs  = path.split('-')
                label = segs[0].replace('%20', ' ')
                val   = '-'.join(segs[1:-1]).replace('%20', ' ').replace('%2F', '/').replace('%2B', '+')
                parts.append(f'{label}: {val}')
            if parts:
                result.append(Paragraph(' | '.join(parts), S['Badge']))
            i += 1
            continue

        # ── H1 ───────────────────────────────────────────────────────────────
        if re.match(r'^# [^#]', line):
            result.append(Paragraph(inline(line[2:].strip()), S['Title']))
            i += 1
            continue

        # ── H2 ───────────────────────────────────────────────────────────────
        if re.match(r'^## [^#]', line):
            result.append(HRFlowable(width='100%', thickness=1.2,
                                     color=C_BLUE, spaceAfter=3, spaceBefore=8))
            result.append(Paragraph(inline(line[3:].strip()), S['H2']))
            i += 1
            continue

        # ── H3 ───────────────────────────────────────────────────────────────
        if re.match(r'^### ', line):
            result.append(Paragraph(inline(line[4:].strip()), S['H3']))
            i += 1
            continue

        # ── Horizontal rule ───────────────────────────────────────────────────
        if re.match(r'^-{3,}\s*$', line) or re.match(r'^\*{3,}\s*$', line):
            result.append(HRFlowable(width='100%', thickness=0.5,
                                     color=C_GRAY, spaceAfter=4, spaceBefore=4))
            i += 1
            continue

        # ── Fenced code block ─────────────────────────────────────────────────
        if line.strip().startswith('```'):
            code_lines = []
            i += 1
            while i < N and not lines[i].strip().startswith('```'):
                code_lines.append(lines[i])
                i += 1
            i += 1  # skip closing ```
            code_text = '\n'.join(code_lines)
            # Preformatted does not do XML parsing – no escaping needed
            result.append(Preformatted(code_text, S['Code']))
            continue

        # ── Markdown table ─────────────────────────────────────────────────
        if line.startswith('|'):
            tbl_lines = []
            while i < N and lines[i].startswith('|'):
                tbl_lines.append(lines[i])
                i += 1
            rows = parse_md_table(tbl_lines)
            if rows:
                t = build_table(rows, S)
                if t:
                    result.append(t)
            continue

        # ── Blockquote ────────────────────────────────────────────────────────
        if line.startswith('> ') or line == '>':
            q_lines = []
            while i < N and (lines[i].startswith('> ') or lines[i] == '>'):
                q_lines.append(lines[i][2:] if lines[i].startswith('> ') else '')
                i += 1

            # Process q_lines – handle nested code blocks
            j, nq, buf = 0, len(q_lines), []
            def flush_buf(buf, result):
                text = ' '.join(l for l in buf if l.strip())
                if text:
                    result.append(Paragraph(inline(text), S['BlockQuote']))
                return []

            while j < nq:
                ql = q_lines[j]
                if ql.strip().startswith('```'):
                    buf = flush_buf(buf, result)
                    code_lines = []
                    j += 1
                    while j < nq and not q_lines[j].strip().startswith('```'):
                        code_lines.append(q_lines[j])
                        j += 1
                    j += 1
                    result.append(Preformatted('\n'.join(code_lines), S['Code']))
                else:
                    buf.append(ql)
                    j += 1
            flush_buf(buf, result)
            continue

        # ── Unordered list ─────────────────────────────────────────────────
        if re.match(r'^(\s{0,3})[-*+] ', line):
            while i < N and re.match(r'^(\s*)[-*+] ', lines[i]):
                m = re.match(r'^(\s*)([-*+]) +(.*)', lines[i])
                if m:
                    depth = len(m.group(1))
                    txt   = inline(m.group(3).strip())
                    st    = S['Bullet2'] if depth >= 2 else S['Bullet']
                    indent_dot = '\xa0\xa0\xa0\xa0•\xa0' if depth >= 2 else '•\xa0'
                    result.append(Paragraph(indent_dot + txt, st))
                i += 1
            continue

        # ── Ordered list ───────────────────────────────────────────────────
        if re.match(r'^\d+\. ', line):
            while i < N and re.match(r'^\d+\. ', lines[i]):
                m = re.match(r'^(\d+)\. +(.*)', lines[i])
                if m:
                    txt = inline(m.group(2).strip())
                    result.append(Paragraph(f'{m.group(1)}.\xa0{txt}', S['Bullet']))
                i += 1
            continue

        # ── Empty line ────────────────────────────────────────────────────
        if not line.strip():
            result.append(Spacer(1, 3))
            i += 1
            continue

        # ── Plain paragraph ───────────────────────────────────────────────
        result.append(Paragraph(inline(line.strip()), S['Normal']))
        i += 1

    return result


# ── Entry point ───────────────────────────────────────────────────────────────
def main():
    base        = r'C:\Claude\Analyze-Windows-Logfiles'
    readme_path = os.path.join(base, 'README.md')
    output_path = os.path.join(base, 'README.pdf')

    with open(readme_path, 'r', encoding='utf-8') as f:
        md_text = f.read()

    S = build_styles()

    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=MARGIN,
        title='Windows Event Analyzer – Dokumentation',
        author='FrankBKW',
        subject='Windows Event Analyzer v1.2.16',
    )

    story = md_to_flowables(md_text, S)

    # Footer separator
    story.append(Spacer(1, 14))
    story.append(HRFlowable(width='100%', thickness=0.5, color=C_GRAY))
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        'Windows Event Analyzer v1.2.16  ·  FrankBKW  ·  2026-05-13',
        S['Footer']))

    doc.build(story)
    print(f'PDF erstellt: {output_path}')


if __name__ == '__main__':
    main()
