import os
import csv
import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor

# Get the directory of the current script
script_dir = os.path.dirname(os.path.realpath(__file__))

# Construct the paths to the dependencies relative to the script's location
kunddata_path = os.path.join(script_dir, 'kunddata.csv')
template_path = os.path.join(script_dir, 'nv4.docx')

# Use these paths in your script
kunddata = kunddata_path
template = template_path

class Kund:
    def __init__(self, kundnr, fornamn, efternamn, adress1, adress21, adress22, email, phone):
        self.kundnr = kundnr
        self.fornamn = fornamn
        self.efternamn = efternamn
        self.adress1 = adress1
        self.adress21 = adress21
        self.adress22 = adress22
        self.email = email
        self.phone = phone

kunddata = kunddata_path # Härifrån hämtar vi kunddata
template = template_path # Här fyller vi i dem
outputDocument = Document(template) # Definiera en output-fil från Word-mallen
kund_list = []  # Kundobjekt lagras här
today = datetime.date.today().strftime("%Y-%m-%d") # Correct way to get today's date


def duplicate_content_on_new_page(document, template): # Duplicera innehållet i mallen  på en ny sida
    document.add_page_break()

    for paragraph in template.paragraphs:
        new_paragraph = document.add_paragraph()
        new_paragraph.style = paragraph.style

        # Kopiera tabbstopp
        for tab_stop in paragraph.paragraph_format.tab_stops:
            new_tab_stop = new_paragraph.paragraph_format.tab_stops.add_tab_stop(tab_stop.position)

        for run in paragraph.runs:
            new_run = new_paragraph.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.size = run.font.size
            new_run.font.name = run.font.name
            new_run.font.color.rgb = run.font.color.rgb
            new_run.font.highlight_color = run.font.highlight_color

            # Kontrollera om texten ser ut som en e-postadress och formatera den som en hyperlänk
            if "@" in run.text:
                new_run.font.color.rgb = RGBColor(0, 0, 255)  # Blå
                new_run.underline = True  # Understruken

    for table in template.tables:
        new_table = document.add_table(rows=0, cols=len(table.columns))
        new_table.style = table.style
        for row in table.rows:
            new_row = new_table.add_row()
            for idx, cell in enumerate(row.cells):
                new_row.cells[idx].text = cell.text

with open(kunddata, mode='r', encoding='utf-8') as csvfile:
    csvreader = csv.reader(csvfile, delimiter='\t')  # Use tab as delimiter
    for row in csvreader:
        kund = Kund(*row)  # Gör en kund av varje rad i CSV-filen
        kund.fornamn = kund.fornamn.title()
        kund.efternamn = kund.efternamn.title()
        kund.adress1 = kund.adress1.title()
        if "lgh" in kund.adress1.lower():
            kund.adress1 = kund.adress1.replace("Lgh", "lgh")
        kund.adress21 = kund.adress21.replace(" ", "")
        if len(kund.adress21) > 3:  
            kund.adress21 = kund.adress21[:3] + " " + kund.adress21[3:] # Lägg till mellanslag efter de första tre siffrorna
        kund.adress22 = kund.adress22.title()
            
        kund_list.append(kund)
        duplicate_content_on_new_page(outputDocument, Document(template))   
        
        for paragraph in outputDocument.paragraphs:  # Loop through all paragraphs in the document
            for run in paragraph.runs:  # Iterate through each run in the paragraph
                placeholders = {"{fornamn}": kund.fornamn, "{efternamn}": kund.efternamn}
                for placeholder, replacement in placeholders.items():
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, replacement)
                        print(f"Replaced {placeholder} with {replacement} in run text: {run.text}")

                if "{adress1}" in run.text:
                    run.text = run.text.replace("{adress1}", kund.adress1)
                if "{adress21}" in run.text:
                    run.text = run.text.replace("{adress21}", kund.adress21)
                if "{adress22}" in run.text:
                    run.text = run.text.replace("{adress22}", kund.adress22)
                if "{today}" in run.text:
                    run.text = run.text.replace("{today}", today)                              


# Spara det uppdaterade dokumentet
outputDocument.save('output.docx')