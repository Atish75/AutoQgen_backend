# AutoQgen_App/services.py

import io
import docx
import pandas as pd
import cloudinary.uploader
from lxml import etree

NAMESPACES = {
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}

def to_pure_lxml(element):
    """BaseOxmlElement ko pure lxml element me convert karta hai"""
    return etree.fromstring(element.xml.encode('utf-8'))

def extract_math(element):
    """OMML math elements ko readable text/latex me nikalta hai."""
    root = to_pure_lxml(element)
    math_elements = root.xpath('.//m:oMath', namespaces=NAMESPACES)
    eqs = []
    for m in math_elements:
        math_texts = m.xpath('.//m:t', namespaces=NAMESPACES)
        eq_text = "".join([t.text for t in math_texts if t.text])
        if eq_text:
            eqs.append(f"${eq_text}$")
    return " | ".join(eqs)

def extract_and_upload_images(element, doc_part):
    """Drawing blips nikal kar Cloudinary par upload karta hai."""
    root = to_pure_lxml(element)
    blips = root.xpath('.//a:blip', namespaces=NAMESPACES)
    image_urls = []
    for blip in blips:
        r_id = blip.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed')
        if r_id and r_id in doc_part.rels:
            rel = doc_part.rels[r_id]
            if "image" in rel.target_ref:
                image_bytes = rel.target_part.blob
                upload_res = cloudinary.uploader.upload(
                    io.BytesIO(image_bytes),
                    folder="exam_diagrams",
                    resource_type="image"
                )
                image_urls.append(upload_res.get("secure_url"))
    return " | ".join(image_urls)

def process_docx_file(uploaded_file):
    """DOCX file se paragraphs aur tables parse karke DataFrame banata hai."""
    doc = docx.Document(io.BytesIO(uploaded_file.read()))
    records = []
    idx = 1

    # 1. Paragraphs scan
    for p in doc.paragraphs:
        text = p.text.strip()
        p_elm = p._element
        eqs = extract_math(p_elm)
        imgs = extract_and_upload_images(p_elm, doc.part)

        if text or eqs or imgs:
            records.append({
                "item_id": idx,
                "type": "paragraph",
                "content": text,
                "equations": eqs,
                "diagram_urls": imgs
            })
            idx += 1

    # 2. Tables scan
    for t_i, table in enumerate(doc.tables):
        for r_i, row in enumerate(table.rows):
            for c_i, cell in enumerate(row.cells):
                c_text = cell.text.strip()
                cell_elm = cell._element
                eqs = extract_math(cell_elm)
                imgs = extract_and_upload_images(cell_elm, doc.part)
                if c_text or eqs or imgs:
                    records.append({
                        "item_id": idx,
                        "type": f"table_{t_i}_r{r_i}_c{c_i}",
                        "content": c_text,
                        "equations": eqs,
                        "diagram_urls": imgs
                    })
                    idx += 1

    return pd.DataFrame(records)

def upload_dataframe_as_excel(df, public_id):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    output.seek(0)

    # invalidate=True Cloudinary CDN cache clear karta hai
    upload_result = cloudinary.uploader.upload(
        output,
        resource_type="raw",
        public_id=f"{public_id}.xlsx",
        overwrite=True,
        invalidate=True
    )
    return upload_result['secure_url'], upload_result['public_id']