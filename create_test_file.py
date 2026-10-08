import docx
from docx.shared import Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from PIL import Image, ImageDraw
import io

doc = docx.Document()

# Document Title
doc.add_heading('Engineering Mathematics & Physics Test', level=1)

# Question 1: Simple Text Question
doc.add_paragraph('Q1: Define Newton\'s Second Law of Motion and state its SI unit.')

# Question 2: Mathematical Equation (OMML Word Math)
p2 = doc.add_paragraph('Q2: Solve the following quadratic equation for x: ')
# Add an actual Word OMML math element (<m:oMath>)
omml_xml = (
    f'<m:oMath {nsdecls("m")}>'
    f'  <m:r><m:t>x = (-b ± √(b² - 4ac)) / 2a</m:t></m:r>'
    f'</m:oMath>'
)
p2._element.append(parse_xml(omml_xml))

# Question 3: Question with an Embedded Diagram/Image
doc.add_paragraph('Q3: In the given circuit diagram below, calculate the equivalent resistance between terminal A and B.')

# Dummy Diagram create karna on-the-fly (Pillow se ek circuit diagram image)
img = Image.new('RGB', (400, 200), color=(240, 240, 240))
draw = ImageDraw.Draw(img)
draw.rectangle([50, 50, 350, 150], outline="black", width=3)
draw.text((120, 90), "CIRCUIT DIAGRAM (R1, R2)", fill="black")

img_bytes = io.BytesIO()
img.save(img_bytes, format='PNG')
img_bytes.seek(0)

# Image ko DOCX me embed karein
doc.add_picture(img_bytes, width=Inches(3.5))

# Question 4: Table Format Question
doc.add_heading('Section B: Numerical Data', level=2)
table = doc.add_table(rows=3, cols=3)
headers = ['Component', 'Value', 'Tolerance']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h

data = [
    ['Resistor R1', '100 Ohm', '±5%'],
    ['Capacitor C1', '10 uF', '±10%']
]

for row_idx, row_data in enumerate(data):
    for col_idx, val in enumerate(row_data):
        table.rows[row_idx + 1].cells[col_idx].text = val

# Save File
filename = "sample_test_paper.docx"
doc.save(filename)
print(f"File successfully created: {filename}")