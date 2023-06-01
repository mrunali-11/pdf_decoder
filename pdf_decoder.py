import PyPDF4
import streamlit as st
import os
import win32com.client
import docx

def read_pdf(file_path):
    with open(file_path, 'rb') as f:
        pdf_reader = PyPDF4.PdfFileReader(f)
        text = ''
        for page in range(pdf_reader.getNumPages()):
            text += pdf_reader.getPage(page).extractText()
    return text

def read_doc(file_path):
    word = win32com.client.Dispatch("Word.Application")
    word.visible = False
    doc = word.Documents.Open(file_path)
    text = doc.Content.Text
    doc.Close()
    word.Quit()
    return text

def read_docx(file_path):
    doc = docx.Document(file_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text
    return text

def main():
    st.title("File Reader")

    uploaded_file = st.file_uploader("Upload a file", type=["pdf","docx", "doc", "txt"])

    if uploaded_file is not None:
        st.write("File uploaded successfully.")
        if uploaded_file.type == 'application/pdf':
            # Get current working directory path
            cwd = os.getcwd()

            # Join filename with current working directory path
            file_path = os.path.join(cwd, uploaded_file.name)

            text = read_pdf(file_path)

        elif uploaded_file.type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            text = read_docx(uploaded_file)
        elif uploaded_file.type == 'application/msword':
            # Get current working directory path
            cwd = os.getcwd()

            # Join filename with current working directory path
            file_path = os.path.join(cwd, uploaded_file.name)

            text = read_doc(file_path)
        elif uploaded_file.type == 'text/plain':
            text = uploaded_file.read().decode('utf-8')
        else:
            st.write("Invalid file format. Please upload a PDF or a TXT file.")
            return

        st.write(text)

if __name__ == "__main__":
    main()