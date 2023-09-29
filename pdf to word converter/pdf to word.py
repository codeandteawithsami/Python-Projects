import fitz #pip install PyMuPDF
from docx import Document
    #pip install python-docx

def pdf_to_word(pdf_path,word_path):
    pdf_doc = fitz.open(pdf_path)
    doc = Document()
    
    for page_num in range(pdf_doc.page_count):
        page = pdf_doc.load_page(page_num)
        text = page.get_text()
        doc.add_paragraph(text)
        
    doc.save(word_path)
    
if __name__ == "__main__":
    pdf_file_path = "Python C+.pdf"
    word_file_path = "word.docs"
    pdf_to_word(pdf_file_path,word_file_path)
    print(f"Pdf Converted to word: {word_file_path}")