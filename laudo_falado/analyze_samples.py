import os
from docx import Document

def get_text(file_path):
    try:
        doc = Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    except Exception as e:
        return f"Error reading {file_path}: {e}"

# Sample files
cat_files = [
    "/data/laudos_base_conhecimento/10 OUTUBRO/PACIENTE/41.124 Edson Silva Amorim/Edson Silva Amorim.docx",
    "/data/laudos_base_conhecimento/10 OUTUBRO/PACIENTE/41.240 Paulo Rodrigues de Oliveira/Paulo Rodrigues de Oliveira.docx",
]

ptca_files = [
    "/data/laudos_base_conhecimento/07 JULHO/PACIENTE/39.579 Arlindo de Almeida PTCA/Arlindo de Almeida PTCA.docx",
    "/data/laudos_base_conhecimento/07 JULHO/PACIENTE/39.885 Sebastiana Alves da Cunha - PTCA/Sebastiana Alves da Cunha-PTCA.docx",
]

print("--- CAT SAMPLES ---")
for f in cat_files:
    print(f"\nFILE: {f}")
    print(get_text(f))
    print("-" * 20)

print("\n--- PTCA SAMPLES ---")
for f in ptca_files:
    print(f"\nFILE: {f}")
    print(get_text(f))
    print("-" * 20)
