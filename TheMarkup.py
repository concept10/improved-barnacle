import xlrd
import xml.etree.ElementTree as ET
from collections import defaultdict
import pytesseract
from PIL import Image

# Open the "Tracker" spreadsheet
workbook = xlrd.open_workbook('Tracker.xlsx')
sheet = workbook.sheet_by_index(0)

# Read the tag names from column A of the "Tracker" spreadsheet
tag_names = [sheet.cell_value(row, 0) for row in range(sheet.nrows)]

# Create a mapping from tag names to unique IDs and RGBA colors
tag_map = defaultdict(lambda: {'id': 0, 'color': [0, 0, 0, 1]})
color_index = 0
for tag in tag_names:
  if tag not in tag_map:
    tag_map[tag]['id'] = color_index
    tag_map[tag]['color'] = [color_index, color_index, color_index, 1]
    color_index += 1

# Open the PMD spreadsheet
pmd_workbook = xlrd.open_workbook('pmd.xlsx')
pmd_sheet = pmd_workbook.sheet_by_index(0)

# Read the tag names from column A of the PMD spreadsheet
pmd_tag_names = [pmd_sheet.cell_value(row, 0) for row in range(pmd_sheet.nrows)]

# Associate the tag names from the PMD spreadsheet with those from the "Tracker" spreadsheet
for pmd_tag in pmd_tag_names:
  if pmd_tag in tag_map:
    tag_info = tag_map[pmd_tag]
    tag_info['count'] = tag_info.get('count', 0) + 1

# Create the root element of the XML schema
root = ET.Element("Markups")
root.set("xmlns", "http://schemas.datacontract.org/2004/07/Bluebeam.PDF.Annotations.Markup")
root.set("xmlns:i", "http://www.w3.org/2001/XMLSchema-instance")

# Search the workDirectory for each PDF file associated with each tag name
workDirectory = 'path/to/workDirectory'
for filename in os.listdir(workDirectory):
  if filename.endswith(".pdf"):
    filepath = os.path.join(workDirectory, filename)

    # Use OCR to extract the location of each tag name in the PDF file
    text = pytesseract.image_to_string(Image.open(filepath))

    # Update the count of each tag name in the tag map
    for tag, tag_info in tag_map.items():
      if tag in text:
        tag_info['count'] = tag_info.get('count', 0) + 1

    # Create an XML element for each tag name in the tag map
    for tag, tag_info in tag_map.items():
      markup = ET.SubElement(root, "Markup")
      markup.set("ID", str(tag_info

xml = xmlbuilder.create('BAX', { version: '1.0', encoding: 'UTF-8' })

for tag, data in tag_data.items():
    xml.ele('Markup')
        .ele('TagName', tag).up()
        .ele('TagID', data['id']).up()
        .ele('Color', data['color']).up()
        .ele('PageNumber', data['page_number']).up()
        .ele('Coordinates', data['coordinates']).up()

xml_str = xml.end({ pretty: True})

with open("output.bax", "w") as f:
    f.write(xml_str)
