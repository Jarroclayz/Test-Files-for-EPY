"""
docx_to_image.py

Extracts tables from a .docx file and renders each table as a PNG image
using matplotlib. The output image is saved alongside the source file as
{original_stem}_ss.png.

Usage:
    python scripts/docx_to_image.py path/to/file.docx [path/to/other.docx ...]
"""

import sys
import os
from pathlib import Path

from docx import Document
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


def _wrap_text(text: str, max_chars: int) -> str:
    """Naively wrap text by inserting newlines every max_chars characters at word boundaries."""
    words = text.split()
    lines = []
    current = []
    length = 0
    for word in words:
        if length + len(word) + (1 if current else 0) > max_chars and current:
            lines.append(" ".join(current))
            current = [word]
            length = len(word)
        else:
            current.append(word)
            length += len(word) + (1 if len(current) > 1 else 0)
    if current:
        lines.append(" ".join(current))
    return "\n".join(lines)


def render_table_as_image(table_data: list, output_path: str) -> None:
    """
    Render a 2-D list of strings (table_data[row][col]) as a PNG image.

    The first row is treated as the header row and is rendered with a
    coloured background and bold text.
    """
    if not table_data:
        print("  [warn] Table is empty — skipping.")
        return

    num_rows = len(table_data)
    num_cols = max(len(row) for row in table_data)

    # Normalise rows so every row has the same number of columns
    for row in table_data:
        while len(row) < num_cols:
            row.append("")

    # Determine column widths based on content (proportional)
    col_max_chars = [1] * num_cols
    for row in table_data:
        for c, cell in enumerate(row):
            col_max_chars[c] = max(col_max_chars[c], len(cell))

    total_chars = sum(col_max_chars)
    col_ratios = [n / total_chars for n in col_max_chars]

    # Wrap cell text so it doesn't stretch the figure too wide
    WRAP_THRESHOLD = 60  # characters per column unit
    wrapped = []
    for row in table_data:
        new_row = []
        for c, cell in enumerate(row):
            max_c = max(15, int(col_ratios[c] * WRAP_THRESHOLD / max(col_ratios) if max(col_ratios) > 0 else 15))
            new_row.append(_wrap_text(cell, max_c))
        wrapped.append(new_row)

    # Figure sizing
    col_widths = [max(1.5, r * 14) for r in col_ratios]  # inches
    fig_width = sum(col_widths) + 0.4
    row_height = 0.55  # base height per row (inches)

    # Increase row height when cells contain wrapped text
    row_heights = []
    for row in wrapped:
        max_lines = max(cell.count("\n") + 1 for cell in row)
        row_heights.append(row_height * max_lines)

    fig_height = sum(row_heights) + 0.3
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.set_xlim(0, fig_width)
    ax.set_ylim(0, fig_height)
    ax.axis("off")

    HEADER_COLOR = "#2C5F8A"
    HEADER_TEXT_COLOR = "white"
    ROW_COLOR_ODD = "#FFFFFF"
    ROW_COLOR_EVEN = "#EAF2FB"
    BORDER_COLOR = "#2C5F8A"
    TEXT_COLOR = "#1A1A1A"

    # Draw cells row by row (top to bottom)
    y_cursor = fig_height - 0.15  # start from top, leaving a small margin

    for r, row in enumerate(wrapped):
        rh = row_heights[r]
        x_cursor = 0.2
        bg_color = HEADER_COLOR if r == 0 else (ROW_COLOR_ODD if r % 2 == 1 else ROW_COLOR_EVEN)
        txt_color = HEADER_TEXT_COLOR if r == 0 else TEXT_COLOR
        font_weight = "bold" if r == 0 else "normal"
        font_size = 10 if r == 0 else 9

        for c, cell_text in enumerate(row):
            cw = col_widths[c]
            rect = plt.Rectangle(
                (x_cursor, y_cursor - rh),
                cw,
                rh,
                linewidth=1.2,
                edgecolor=BORDER_COLOR,
                facecolor=bg_color,
                zorder=2,
            )
            ax.add_patch(rect)
            ax.text(
                x_cursor + cw / 2,
                y_cursor - rh / 2,
                cell_text,
                ha="center",
                va="center",
                fontsize=font_size,
                fontweight=font_weight,
                color=txt_color,
                wrap=False,
                zorder=3,
            )
            x_cursor += cw

        y_cursor -= rh

    plt.tight_layout(pad=0)
    plt.savefig(output_path, dpi=150, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    print(f"  [ok] Saved image → {output_path}")


def convert_docx(docx_path: str) -> None:
    """Convert all tables in a .docx file to PNG screenshots."""
    p = Path(docx_path)
    if not p.exists():
        print(f"[error] File not found: {docx_path}")
        return

    doc = Document(str(p))
    tables = doc.tables

    if not tables:
        print(f"[warn] No tables found in {p.name} — skipping.")
        return

    print(f"[info] Processing '{p.name}' — {len(tables)} table(s) found.")

    if len(tables) == 1:
        output_path = p.parent / f"{p.stem}_ss.png"
        table_data = [[cell.text for cell in row.cells] for row in tables[0].rows]
        render_table_as_image(table_data, str(output_path))
    else:
        for i, table in enumerate(tables, start=1):
            output_path = p.parent / f"{p.stem}_ss_{i}.png"
            table_data = [[cell.text for cell in row.cells] for row in table.rows]
            render_table_as_image(table_data, str(output_path))


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python scripts/docx_to_image.py <file.docx> [<file2.docx> ...]")
        sys.exit(1)

    for path in sys.argv[1:]:
        convert_docx(path)
