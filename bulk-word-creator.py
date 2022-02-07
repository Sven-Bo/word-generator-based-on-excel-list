from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

word_template_path = Path(__file__).parent / "vendor-contract.docx"
excel_path = Path(__file__).parent / "contracts-list.xlsx"
output_dir = Path(__file__).parent / "OUTPUT"

output_dir.mkdir(exist_ok=True)

df = pd.read_excel(excel_path)
df["TODAY"] = pd.to_datetime(df["TODAY"]).dt.date
df["TODAY_IN_ONE_WEEK"] = pd.to_datetime(df["TODAY_IN_ONE_WEEK"]).dt.date

for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['VENDOR']}-contract.docx"
    doc.save(output_path)
