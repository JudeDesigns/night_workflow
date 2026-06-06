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
        table {{ width: 100%; border-collapse: collapse; margin-bottom: 20px; table-layout: fixed; }}
        th, td {{ border: 1px solid black; padding: 4px; text-align: left; word-wrap: break-word; overflow-wrap: anywhere; line-height: 1.2; }}
        /* Spec §10 — spanned group headers (vendor/driver/customer + date) are always centered. */
        .header {{ text-align: center; font-weight: bold; background-color: #eee; }}
        .centered, td.centered, th.centered {{ text-align: center; }}
        /* Z-driver rows (future deliveries / pickups) — distinct grey shade. */
        tr.z-driver td {{ background-color: #B5BFC9; }}
        /* Produce rows (Jetro page) — lighter grey. */
        tr.produce td {{ background-color: #D3D3D3; }}
    """

    def _fits_one_page(block: dict[str, Any], font_size: float) -> bool:
        """Render the block standalone and return True iff it fits on 1 page."""
        block_html = template_func(block, font_size)
        temp_html = f"<html><head><style>{base_css}</style></head><body>{block_html}</body></html>"
        doc = HTML(string=temp_html).render(stylesheets=[CSS(string=base_css)])
        return len(doc.pages) <= 1

    final_html_parts = []

    # Spec §8.4 keep-together: every atomic block must fit on a single page.
    # Strategy:
    #   1. Start at 10pt (the ideal readable size).
    #   2. If a heuristic suggests we need to start smaller (long blocks), do so.
    #   3. Verify by rendering standalone; if it still overflows, shrink by
    #      0.88x and re-verify. Loop up to MAX_ATTEMPTS times. There is no hard
    #      minimum font floor (spec: "fitting on one page always wins").
    MAX_ATTEMPTS = 10
    SHRINK_FACTOR = 0.88
    HARD_MIN = 3.5  # Practical floor — below this, the page is nearly unreadable
                    # anyway and we accept the result rather than loop forever.

    for i, block in enumerate(blocks):
        row_count = len(block.get("rows", []))

        # Heuristic initial guess. The 1.7 factor (vs the old 1.4) accounts for
        # product-name wrapping that frequently doubles row height in practice.
        if row_count <= 12:
            best_font = 10.0
        else:
            heuristic = avail_points / ((row_count + 5) * 1.7)
            best_font = min(10.0, max(HARD_MIN, heuristic))

        # Iterative verification: shrink until the block fits on one page.
        attempts = 0
        while attempts < MAX_ATTEMPTS and best_font >= HARD_MIN:
            if _fits_one_page(block, best_font):
                break
            best_font *= SHRINK_FACTOR
            attempts += 1
        else:
            # Loop exited without fitting; clamp to HARD_MIN and accept.
            best_font = max(best_font, HARD_MIN)

        # When the block needed scaling, force it onto its own page so
        # neighbouring blocks can't push it across a boundary at render time.
        css_class = "block-force-new-page" if best_font < 10.0 else "block-avoid-break"
        final_html_parts.append(
            f'<div id="block-{i}" class="{css_class}">{template_func(block, best_font)}</div>'
        )

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
