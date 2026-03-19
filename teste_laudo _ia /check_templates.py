from docx import Document

def get_text(file_path):
    try:
        doc = Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"Error reading {file_path}: {e}"

templates = ["Laudo de cateterismo.docx", "Laudo de Angioplastia.docx"]

for t in templates:
    print(f"\nTEMPLATE: {t}")
    print(get_text(t))
    print("-" * 20)
