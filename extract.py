import pdfplumber, sys, os

pdfs = [
    "，然=一淼(1).pdf",
    "clas_mech.pdf",
    "普通班理论力学平时大作业.pdf",
]

out_dir = "extracted"
os.makedirs(out_dir, exist_ok=True)

for name in pdfs:
    path = os.path.join(".", name)
    out_name = os.path.splitext(name)[0].replace("，", "").replace(" ", "") + ".txt"
    out_path = os.path.join(out_dir, out_name)
    print(f"=== Processing {name} ===")
    try:
        with pdfplumber.open(path) as pdf:
            n = len(pdf.pages)
            print(f"  pages: {n}")
            with open(out_path, "w", encoding="utf-8") as f:
                for i, page in enumerate(pdf.pages):
                    txt = page.extract_text() or ""
                    f.write(f"\n----- PAGE {i+1} -----\n")
                    f.write(txt)
                    f.write("\n")
        print(f"  -> wrote {out_path}")
    except Exception as e:
        print(f"  ERROR: {e}")
