#!/usr/bin/env python3
"""Generate PDF from CV HTML using WeasyPrint."""

import re
from pathlib import Path

# ── Read source HTML ──────────────────────────────────────────────────────────
src = Path("ammd-cv.html").read_text(encoding="utf-8")

# ── Photo-placeholder CSS ─────────────────────────────────────────────────────
PHOTO_CSS = """
  @page {
    size: A4;
    margin: 0;
  }

  /* Override web layout for print */
  html { font-size: 13px; }

  .page { max-width: 100%; }

  body, .header, .main { background: white; }
  .sidebar { background: #ede9e3; }
  .header { padding: 28px 40px 24px; }
  .sidebar { padding: 28px 18px 28px 40px; }
  .main { padding: 24px 40px 24px 32px; }
  .header::after { display: none; }

  /* Photo placeholder layout */
  .header-inner {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    gap: 20px;
  }

  .header-text { flex: 1; }

  .photo-placeholder {
    width: 108px;
    height: 130px;
    border: 2px dashed #DDD8D1;
    border-radius: 4px;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    background: #F7F4F0;
    color: #8A8480;
    font-size: 0.7rem;
    text-align: center;
    line-height: 1.4;
    padding: 8px;
  }

  .photo-placeholder svg {
    width: 32px;
    height: 32px;
    margin-bottom: 6px;
    opacity: 0.45;
  }

  /* Avoid page breaks inside jobs/sections */
  .job { page-break-inside: avoid; }
  .section { page-break-inside: avoid; }

  /* Show full URLs when printing */
  a[href^="http"]::after {
    content: " (" attr(href) ")";
    font-size: 0.65rem;
    color: #666;
    font-weight: 300;
  }
  .proj-link::after,
  .client-link::after { content: none; }

  /* Remove underlines from links in print */
  a { border-bottom: none !important; }
"""

# ── Inject extra CSS before </style> ─────────────────────────────────────────
src = src.replace("</style>", PHOTO_CSS + "\n</style>")

# ── Wrap header content in .header-inner + add photo placeholder ──────────────
PHOTO_SVG = """<svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
        <circle cx="12" cy="8" r="4" stroke="#8A8480" stroke-width="1.5"/>
        <path d="M4 20c0-4 3.6-7 8-7s8 3 8 7" stroke="#8A8480" stroke-width="1.5" stroke-linecap="round"/>
      </svg>"""

PHOTO_BLOCK = f"""
      <div class="photo-placeholder">
        {PHOTO_SVG}
        <span>Photo<br>3×3.5 cm</span>
      </div>"""

# Wrap the existing header children inside .header-inner
# Find the header opening tag and wrap its contents
src = src.replace(
    '<header class="header">',
    '<header class="header">\n    <div class="header-inner">\n    <div class="header-text">'
)
src = src.replace(
    '</header>',
    f'    </div>{PHOTO_BLOCK}\n    </div>\n  </header>'
)

# ── Write modified HTML ───────────────────────────────────────────────────────
out_html = Path("ammd-cv-pdf.html")
out_html.write_text(src, encoding="utf-8")
print(f"Wrote: {out_html}")

# ── Convert to PDF ────────────────────────────────────────────────────────────
import weasyprint

print("Converting to PDF (this may take a moment)...")
pdf_path = Path("Andrija_Despotovic_CV.pdf")

html = weasyprint.HTML(filename=str(out_html), base_url=".")
css = weasyprint.CSS(string="")  # additional CSS if needed
html.write_pdf(str(pdf_path))

print(f"Done: {pdf_path} ({pdf_path.stat().st_size // 1024} KB)")
