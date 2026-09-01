import pdfplumber, os

os.makedirs("rendered", exist_ok=True)
path = "，然=一淼(1).pdf"
with pdfplumber.open(path) as pdf:
    n = len(pdf.pages)
    print(f"pages: {n}")
    for i, page in enumerate(pdf.pages):
        im = page.to_image(resolution=150)
        out = os.path.join("rendered", f"ran_p{i+1}.png")
        im.save(out)
        print(f"saved {out}")
