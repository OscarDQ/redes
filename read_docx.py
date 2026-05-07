import zipfile
import xml.etree.ElementTree as ET
import os

def extract_docx(file_path):
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        return
    with zipfile.ZipFile(file_path) as docx:
        xml_content = docx.read('word/document.xml')
        
    tree = ET.fromstring(xml_content)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    print("--- TABLES ---")
    for table in tree.findall('.//w:tbl', ns):
        for row in table.findall('.//w:tr', ns):
            row_data = []
            for cell in row.findall('.//w:tc', ns):
                texts = cell.findall('.//w:t', ns)
                cell_text = "".join([t.text for t in texts if t.text])
                row_data.append(cell_text)
            print(" | ".join(row_data))
        print("--- END TABLE ---")

if __name__ == '__main__':
    extract_docx('RSS_Cuadro_Doble_Entrada.docx')
