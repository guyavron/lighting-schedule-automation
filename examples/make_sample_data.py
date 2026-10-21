#!/usr/bin/env python3
"""
Generate all the FICTIONAL input data the demo needs, from code:

  examples/out/legend_template.xlsx          generic empty legend (make_template.py)
  examples/out/photos/*.jpg                  product images drawn in code (draw_products.py)
  examples/out/old_projects/*.xlsx           3 "old project" legends, built with the real
                                             insert_fixture.py -- the raw material for the bank
  examples/spec-sheets/*.pdf                 manufacturer-style spec sheets for invented products,
                                             plus broken ones for the failure tests

Nothing here comes from a real studio, manufacturer or project.
"""
import datetime
import io
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "skill" / "scripts"))
sys.path.insert(0, str(HERE))

import insert_fixture as IF          # noqa: E402
import set_project_header as SPH     # noqa: E402
import draw_products                 # noqa: E402
import make_template                 # noqa: E402

DATA = json.loads((HERE / "data" / "fixtures.json").read_text(encoding="utf8"))
OUT = HERE / "out"
SPECS = HERE / "spec-sheets"


def build_photos():
    d = OUT / "photos"
    d.mkdir(parents=True, exist_ok=True)
    for key, f in DATA["fixtures"].items():
        draw_products.draw(f["img"], f["גוון"], d / f"{key}.jpg")
    draw_products.draw("pendant", "לבן מט", d / "spec_lune_pendant.jpg")
    draw_products.draw("trackspot", "שחור", d / "spec_track_one_15.jpg")
    draw_products.draw("grazer", "אפור גרפיט", d / "spec_riva_120.jpg")
    return d


def build_old_projects(template, photos):
    d = OUT / "old_projects"
    d.mkdir(parents=True, exist_ok=True)
    built = []
    for proj in DATA["old_projects"]:
        path = d / proj["file"]
        shutil.copy(template, path)
        SPH.set_project_header(str(path), proj["name"], proj["color"],
                               build_date=datetime.date.fromisoformat(proj["date"]))
        for key, mark in proj["rows"]:
            f = DATA["fixtures"][key]
            data = {k: f[k] for k in ("תיאור", "יצרן", "מתח_עמעום", "גוון")}
            data["סימון"] = mark
            IF.insert_fixture(str(path), f["cat"], data, image_path=str(photos / f"{key}.jpg"))
        built.append(path)
    return built


# ---------------------------------------------------------------- spec sheets
def _spec_pdf(path, brand, title, subtitle, photo, rows, footer):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setTitle(f"{brand} {title} - spec sheet (fictional)")
    c.setAuthor("fictional demo data")
    c.setCreator("lighting-schedule-automation/examples")
    if brand:
        c.setFont("Helvetica-Bold", 22); c.drawString(20 * mm, h - 25 * mm, brand)
    c.setFont("Helvetica", 9); c.drawRightString(w - 20 * mm, h - 25 * mm, "FICTIONAL PRODUCT - DEMO DATA")
    c.setFont("Helvetica-Bold", 16); c.drawString(20 * mm, h - 40 * mm, title)
    c.setFont("Helvetica", 11); c.drawString(20 * mm, h - 47 * mm, subtitle)
    c.drawImage(str(photo), 20 * mm, h - 135 * mm, width=95 * mm, height=79 * mm)
    y = h - 145 * mm
    c.setFont("Helvetica-Bold", 11); c.drawString(20 * mm, y, "Technical data"); y -= 7 * mm
    for k, v in rows:
        c.setFont("Helvetica", 9.5); c.drawString(20 * mm, y, k); c.drawString(75 * mm, y, v)
        c.setStrokeColorRGB(.85, .85, .85); c.line(20 * mm, y - 2 * mm, w - 20 * mm, y - 2 * mm)
        y -= 6.5 * mm
    c.setFont("Helvetica", 7.5); c.drawString(20 * mm, 15 * mm, footer)
    c.showPage(); c.save()


def build_specs(photos):
    SPECS.mkdir(parents=True, exist_ok=True)
    lune = SPECS / "oskari-lux_lune-pendant.pdf"
    _spec_pdf(lune, "OSKARI LUX", "Lune Pendant", "Suspended decorative luminaire - PENDANT",
              photos / "spec_lune_pendant.jpg", [
                  ("Installation", "Suspension, ceiling rose for recessed mounting"),
                  ("Light source", "LED 9W"), ("Luminous flux", "810 lm"),
                  ("Colour temperature", "2700K - other K, please consult"),
                  ("CRI", ">90"), ("Beam angle", "38° / 72° (choose optic)"),
                  ("Power supply", "230V"), ("Dimming", "TRIAC - DALI on request"),
                  ("Finish", "Matt white / Black / Brushed brass"),
                  ("Protection", "IP20"),
              ], "OSKARI LUX is an invented brand. This sheet was generated for a demo repository.")
    track = SPECS / "corvella_track-one-15.pdf"
    _spec_pdf(track, "CORVELLA", "Track One 15", "Spotlight for Linea 48 magnetic track",
              photos / "spec_track_one_15.jpg", [
                  ("Installation", "Track-mounted, Linea 48 system"),
                  ("Light source", "LED 12W"), ("Luminous flux", "1150 lm"),
                  ("Colour temperature", "3000K"), ("CRI", "95"),
                  ("Beam angle", "24°"), ("Power supply", "48V"), ("Dimming", "DALI"),
                  ("Finish", "Black"),
              ], "CORVELLA is an invented brand. This sheet was generated for a demo repository.")
    riva = SPECS / "tobek_riva-120.pdf"
    _spec_pdf(riva, "TOBEK", "Riva 120", "Linear wall grazer, indoor / outdoor",
              photos / "spec_riva_120.jpg", [
                  ("Installation", "Wall-mounted"),
                  ("Light source", "LED 14W"), ("Luminous flux", "1250 lm"),
                  ("Colour temperature", "3000K"), ("CRI", "90"),
                  ("Beam angle", "10° x 60° (asymmetric)"), ("Protection", "IP65"),
                  ("Power supply", "24V"), ("Dimming", "DALI"),
                  ("Finish", "Anthracite / White"),
              ], "TOBEK is an invented brand. This sheet was generated for a demo repository.")
    # failure cases ---------------------------------------------------------
    bad = SPECS / "broken"
    bad.mkdir(exist_ok=True)
    # 1) scanned: the same sheet as a picture only, no text layer
    import subprocess
    from PIL import Image
    tmp = OUT / "_scan"
    tmp.mkdir(parents=True, exist_ok=True)
    subprocess.run(["pdftoppm", "-jpeg", "-r", "70", "-singlefile", str(lune), str(tmp / "p")], check=True)
    Image.open(tmp / "p.jpg").convert("RGB").save(bad / "scanned_no-text-layer.pdf", "PDF", resolution=70)
    # 2) truncated download
    raw = lune.read_bytes()
    (bad / "truncated.pdf").write_bytes(raw[: len(raw) // 3])
    # 3) not a PDF at all (e.g. a saved web page renamed to .pdf)
    (bad / "not-really-a.pdf").write_text("<html><body>Product page</body></html>", encoding="utf8")
    # 4) a generic OEM sheet with no brand name anywhere
    _spec_pdf(bad / "no-brand_generic-downlight.pdf", "", "Downlight 75", "Recessed downlight",
              photos / "kito.jpg", [("Light source", "LED 8W"), ("Luminous flux", "720 lm"),
                                    ("Colour temperature", "3000K"), ("Power supply", "350mA")],
              "Generic sheet without manufacturer details (demo).")
    return [lune, track, riva]


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    template = make_template.build(OUT / "legend_template.xlsx")
    photos = build_photos()
    old = build_old_projects(template, photos)
    specs = build_specs(photos)
    print(f"template: {template}\nold projects: {len(old)}\nspec sheets: {[s.name for s in specs]}")


if __name__ == "__main__":
    main()
