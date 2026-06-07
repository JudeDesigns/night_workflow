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

    def _fits_pages(block: dict[str, Any], font_size: float, max_pages: int) -> bool:
        """Render the block standalone and return True iff it fits in max_pages."""
        block_html = template_func(block, font_size)
        temp_html = f"<html><head><style>{base_css}</style></head><body>{block_html}</body></html>"
        doc = HTML(string=temp_html).render(stylesheets=[CSS(string=base_css)])
        return len(doc.pages) <= max_pages

    final_html_parts = []

    # Fitting strategy:
    #   Phase 1 — Try to fit on a single page at a readable font (>= READABLE_MIN).
    #             Start at 10pt and shrink by 0.88x per attempt.
    #   Phase 2 — If a single page would require shrinking below READABLE_MIN,
    #             allow the block to span 2 pages instead. Restart at 10pt, shrink
    #             only if 2 pages still doesn't fit. The spanning block claims
    #             both pages exclusively (page-break-after: always on the wrapper
    #             prevents the next block from sharing the second page) and rows
    #             never split mid-row (page-break-inside: avoid on every <tr>).
    MAX_ATTEMPTS = 10
    SHRINK_FACTOR = 0.88
    READABLE_MIN = 8.0  # below this, single-page mode is rejected in favor of 2-page span
    HARD_MIN = 3.5      # absolute floor for the 2-page attempt

    for i, block in enumerate(blocks):
        row_count = len(block.get("rows", []))

        # Heuristic initial guess for Phase 1. The 1.7 factor accounts for
        # product-name wrapping that often doubles row height in practice.
        if row_count <= 12:
            best_font = 10.0
        else:
            heuristic = avail_points / ((row_count + 5) * 1.7)
            best_font = min(10.0, max(READABLE_MIN, heuristic))

        # Phase 1 — fit on one page at >= READABLE_MIN.
        fits_one = False
        attempts = 0
        while attempts < MAX_ATTEMPTS and best_font >= READABLE_MIN:
            if _fits_pages(block, best_font, 1):
                fits_one = True
                break
            best_font *= SHRINK_FACTOR
            attempts += 1

        if fits_one:
            span_pages = 1
        else:
            # Phase 2 — allow 2 pages. Reset to 10pt (no need to shrink yet).
            span_pages = 2
            best_font = 10.0
            attempts = 0
            while attempts < MAX_ATTEMPTS and best_font >= HARD_MIN:
                if _fits_pages(block, best_font, 2):
                    break
                best_font *= SHRINK_FACTOR
                attempts += 1
            else:
                best_font = max(best_font, HARD_MIN)

        # Wrapper class selection:
        #   - 2-page span: claim both pages exclusively (break before AND after).
        #     Do NOT set page-break-inside: avoid — the block must be allowed
        #     to break across the two pages.
        #   - 1-page but shrunk: force a fresh page so neighbours can't push it.
        #   - 1-page at 10pt: just avoid mid-block breaks.
        if span_pages == 2:
            css_class = "block-span-pages"
        elif best_font < 10.0:
            css_class = "block-force-new-page"
        else:
            css_class = "block-avoid-break"
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
            /* A block that intentionally spans 2 pages: claim both pages
               exclusively via break-before + break-after. Crucially, do NOT
               set page-break-inside: avoid here — the block must break. */
            .block-span-pages {{ page-break-before: always; page-break-after: always; }}
            /* Never split a row across pages — when a block spans pages, the
               break always lands between rows. */
            tr {{ page-break-inside: avoid; }}
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
