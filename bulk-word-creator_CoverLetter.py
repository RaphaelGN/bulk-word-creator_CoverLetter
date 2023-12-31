from pathlib import Path

import pandas as pd  # pip install pandas openpyxl
from docxtpl import DocxTemplate  # pip install docxtpl

base_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
word_template_path = base_dir / "cover_letter.docx"
excel_path = base_dir / "cv_classeur.xlsx"
output_dir = base_dir / "OUTPUT"

# Create output folder for the word documents
output_dir.mkdir(exist_ok=True)

# Convert Excel sheet to pandas dataframe
df = pd.read_excel(excel_path, sheet_name="Feuil1")

# Keep only date part YYYY-MM-DD (not the time)

from datetime import date, datetime
import arrow

today = date.today()
d2 = today.strftime("%B ")
d3=(arrow.get(datetime.utcnow()).format('Do'))
d4=d2+d3+today.strftime(", %Y")
df["date"]=d4

now = datetime.now()

d5=now.strftime("%H:%M:%S")
df["heure"]=d5


#pd.to_datetime(df["TODAY"]).dt.date

#df["TODAY_IN_ONE_WEEK"] = pd.to_datetime(df["TODAY_IN_ONE_WEEK"]).dt.date

#Iterate over each row in df and render word document
from docx2pdf import convert

for record in df.to_dict(orient="records"):
    doc = DocxTemplate(word_template_path)
    doc.render(record)
    output_path = output_dir / f"{record['entreprise']}-raphael.docx"
    
    doc.save(output_path)
    convert(output_dir / f"{record['entreprise']}-raphael.docx", output_dir / f"{record['entreprise']}-raphael.pdf")

