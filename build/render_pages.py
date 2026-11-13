import sys
from pathlib import Path
import pypdfium2 as pdfium

pdf_path = Path(sys.argv[1])
out_dir = Path(sys.argv[2])
pages = [int(x) for x in sys.argv[3].split(',')] if len(sys.argv) > 3 else None
scale = 2.0
out_dir.mkdir(parents=True, exist_ok=True)

pdf = pdfium.PdfDocument(str(pdf_path))
total = len(pdf)
print(f'pages={total}')
targets = pages if pages else range(1, total + 1)
for pno in targets:
    if pno < 1 or pno > total:
        continue
    page = pdf[pno - 1]
    img = page.render(scale=scale).to_pil()
    out = out_dir / f'page-{pno:02d}.png'
    img.save(out)
    print(out)
