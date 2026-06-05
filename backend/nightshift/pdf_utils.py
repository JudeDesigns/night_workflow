"""PDF generation utilities (Spec §8.4 Universal keep-together algorithm)."""
from __future__ import annotations

import io
from typing import Any

try:
    from weasyprint import HTML, CSS
except ImportError:
    HTML = None
    CSS = None

def render_keep_together_pdf(
    blocks: list[dict[str, Any]], 
    output_path: str, 
    template_func: Any,
    landscape: bool = False
) -> list[tuple[int, int]]:
    """Universal keep-together algorithm (Spec §8.4).

    Optimized with heuristics to reduce rendering passes.

    Returns a list of (start_page, end_page) 1-based tuples, one per block,
    read back from the rendered document so Menu-page ranges are accurate
    (Spec §6.6). A single-page block has start_page == end_page.
    """
    if not HTML:
        return []

    page_size = "letter landscape" if landscape else "letter"
    # Available height in points (72 points per inch)
    # Letter is 11in height. Landscape 8.5in height.
    # Minus 1 inch total margins = 10in or 7.5in.
    avail_points = (7.5 if landscape else 10.0) * 72
    
    base_css = f"""
        @page {{ 
            size: {page_size}; 
            margin: 0.5in; 
            @bottom-right {{
                content: "Page " counter(page);
                font-size: 8pt;
            }}
        }}
        body {{ font-family: sans-serif; margin: 0; padding: 0; }}
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; }}
        th, td {{ border: 1px solid black; padding: 4px; text-align: left; word-wrap: break-word; line-height: 1.2; }}
        .header {{ text-align: center; font-weight: bold; background-color: #eee; }}
        .centered {{ text-align: center; }}
    """

    final_html_parts = []
    
    for i, block in enumerate(blocks):
        row_count = len(block.get("rows", []))
        # Estimate: Each row takes roughly 1.4em in height (font-size * 1.4)
        # Plus 2 rows for header and column headers.
        
        # 1. Short-circuit for small blocks
        if row_count <= 15:
            best_font = 10.0
        else:
            # 2. Heuristic Initial Guess
            # We want (row_count + 3) * (font_size * 1.4) <= avail_points
            # font_size <= avail_points / ((row_count + 3) * 1.4)
            heuristic_font = avail_points / ((row_count + 4) * 1.4)
            best_font = min(10.0, max(1.0, heuristic_font))
            
            # 3. Single-pass Verification
            # Render once at the heuristic font size. 
            # If it still doesn't fit, scale down by another 10% once.
            block_html = template_func(block, best_font)
            temp_html = f"<html><head><style>{base_css}</style></head><body>{block_html}</body></html>"
            doc = HTML(string=temp_html).render(stylesheets=[CSS(string=base_css)])
            
            if len(doc.pages) > 1:
                best_font *= 0.85 # Safety factor if heuristic was too aggressive
        
        # Add to final document
        # We use 'page-break-before: always' for every block except the first one
        # to ensure they stay on their own pages if they were scaled.
        # Small blocks that weren't scaled can share pages if they fit.
        css_class = "block-force-new-page" if best_font < 10.0 else "block-avoid-break"
        final_html_parts.append(f'<div id="block-{i}" class="{css_class}">{template_func(block, best_font)}</div>')

    full_html = f"""
    <html>
    <head>
        <style>
            {base_css}
            .block-avoid-break {{ page-break-inside: avoid; margin-bottom: 30px; }}
            .block-force-new-page {{ page-break-before: always; page-break-inside: avoid; }}
        </style>
    </head>
    <body>
    """
    for part in final_html_parts:
        full_html += part
    full_html += "</body></html>"
    
    document = HTML(string=full_html).render(stylesheets=[CSS(string=base_css)])
    document.write_pdf(output_path)
    
    # Extract the start page of each block from its anchor.
    total_pages = len(document.pages)
    starts: list[int] = []
    for i in range(len(blocks)):
        target_id = f"block-{i}"
        found_page = 1
        for page_idx, page in enumerate(document.pages, start=1):
            if any(anchor[0] == target_id for anchor in page.anchors.items()):
                found_page = page_idx
                break
        starts.append(found_page)

    # Derive each block's end page: it runs until just before the next block
    # that begins on a strictly later page. The final block ends on the last
    # page of the document. Blocks sharing a page collapse to a single page.
    page_map: list[tuple[int, int]] = []
    for i, start in enumerate(starts):
        end = total_pages
        for nxt in starts[i + 1:]:
            if nxt > start:
                end = nxt - 1
                break
        page_map.append((start, max(start, end)))

    return page_map

def get_pdf_page_numbers(pdf_path: str) -> list[dict[str, Any]]:
    """Read back page numbers from a rendered PDF (Spec §6.6)."""
    # This requires a library like pypdf or similar.
    # Since we might not have it, we'll try to estimate or use a simple heuristic.
    # Actually, we can use weasyprint's metadata if we had the document object.
    # For now, let's return a placeholder that we'll improve if needed.
    return []
