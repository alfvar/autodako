import csv
from datetime import datetime
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT

# Ange sökvägen till din CSV-fil och Word-mall
kunddata = 'kunddata.csv'
template = 'nv4.docx'

# Ladda Word-mallen
doc = Document(template)

# Funktion för att duplicera innehållet i en Word-mall och lägga till det på en ny sida
def duplicate_content_on_new_page(document, template):
    document.add_page_break()

    for paragraph in template.paragraphs:
        new_paragraph = document.add_paragraph()
        new_paragraph.style = paragraph.style

        # Kopiera tabbstopp
        for tab_stop in paragraph.paragraph_format.tab_stops:
            new_tab_stop = new_paragraph.paragraph_format.tab_stops.add_tab_stop(tab_stop.position)
            # Här kan du också kopiera andra egenskaper av tab_stop om så behövs, som ledartyp och justering

        for run in paragraph.runs:
            new_run = new_paragraph.add_run(run.text)
            new_run.bold = run.bold
            new_run.italic = run.italic
            new_run.underline = run.underline
            new_run.font.size = run.font.size
            new_run.font.name = run.font.name
            new_run.font.color.rgb = run.font.color.rgb
            new_run.font.highlight_color = run.font.highlight_color
            # Fortsätt att kopiera annan formatering efter behov...

    for table in template.tables:
        new_table = document.add_table(rows=0, cols=len(table.columns))
        new_table.style = table.style
        for row in table.rows:
            new_row = new_table.add_row()
            for idx, cell in enumerate(row.cells):
                new_row.cells[idx].text = cell.text
                # Kopiera cellformatering om så behövs

# Öppna CSV-filen och läs in datan
with open(kunddata, newline='', encoding='utf-8') as csvfile:
	customerdata = csv.reader(csvfile, delimiter='\t')
	for row in customerdata:
		try:
			# Tilldela varje fält i raden till en variabel
			kundnr, fornamn, efternamn, adress1, adress21, adress22, email, phone = row
		except ValueError:
			print(f"Fel format på rad: {row}")
			continue  # Hoppa över denna rad och fortsätt med nästa

		# Duplicera innehållet i Word-mallen på en ny sida för varje kund
		duplicate_content_on_new_page(doc, Document(template))

		# Här kan du lägga till kod för att anpassa det duplicerade innehållet
		# baserat på kunddata, t.ex. ersätta platshållartext med faktiska värden

# Spara det uppdaterade dokumentet
doc.save('output.docx')