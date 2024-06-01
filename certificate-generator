import pandas as pd
from datetime import date
from docxtpl import DocxTemplate
from tkinter import Tk, filedialog
from docx2pdf import convert

def preencher_documento_template(hj, template, nome, serie, escola, matricula, quantidade, modelo, imei, tb, id):
    """
    Fills the document template with the provided data.
    
    Parameters:
    hj (str): Today's date.
    template (str): Path to the template file.
    nome (str): Name of the student.
    serie (str): Series of the student.
    escola (str): School of the student.
    matricula (str): Student's enrollment number.
    quantidade (int): Quantity.
    modelo (str): Model.
    imei (str): IMEI number.
    tb (str): Tomb.
    id (str): ID.
    
    Returns:
    doc (DocxTemplate): Rendered document.
    """
    doc = DocxTemplate(template)
    context = {'nome': nome, 'serie': serie, 'escola': escola, 'matricula': matricula, 'quantidade': quantidade, 'modelo': modelo, 'imei': imei, 'tombo': tb, 'id': id, 'data': hj}
    doc.render(context)
    return doc

def selecionar_arquivo_planilha():
    """
    Opens a file dialog to select the spreadsheet file.
    
    Returns:
    str: Path to the selected spreadsheet file.
    """
    root = Tk()
    root.withdraw()  # Hide the main window
    arquivo_planilha = filedialog.askopenfilename(title="Select Spreadsheet")
    return arquivo_planilha

def selecionar_arquivo_template():
    """
    Opens a file dialog to select the document template file.
    
    Returns:
    str: Path to the selected document template file.
    """
    root = Tk()
    root.withdraw()  # Hide the main window
    arquivo_template = filedialog.askopenfilename(title="Select Document Template")
    return arquivo_template

def selecionar_diretorio_salvar():
    """
    Opens a file dialog to select the directory to save the new documents.
    
    Returns:
    str: Path to the selected directory.
    """
    root = Tk()
    root.withdraw()  # Hide the main window
    return filedialog.askdirectory(title="Select Directory")

def main():
    """
    Main function to process the spreadsheet and generate documents based on the template.
    """
    # Select the spreadsheet file
    arquivo_planilha = selecionar_arquivo_planilha()

    # Load data from the spreadsheet
    df = pd.read_excel(arquivo_planilha)

    # Select the document template file
    arquivo_template = selecionar_arquivo_template()

    # Select the directory where the new documents will be saved
    diretorio_destino = selecionar_diretorio_salvar()

    # Store the current date
    hoje = date.today().strftime("%d/%m/%Y")

    # Iterate over the data and fill the document for each row
    for index, row in df.iterrows():
        nome = row['nome']
        serie = row['serie']
        escola = row['escola ']
        matricula = row['matricula']
        quantidade = row['qtd']
        modelo = row['modelo']
        imei = row['imei']
        tb = row['tombo']
        id = row['icid']

        novo_documento = preencher_documento_template(hoje, arquivo_template, nome, serie, escola, matricula, quantidade, modelo, imei, tb, id)

        # Save the filled document with the current row data
        nome_arquivo_docx = f"{diretorio_destino}/{nome}.docx"
        novo_documento.save(nome_arquivo_docx)

        # Convert the DOCX document to PDF
        nome_arquivo_pdf = f"{diretorio_destino}/{nome}.pdf"
        convert(nome_arquivo_docx, nome_arquivo_pdf)

if __name__ == "__main__":
    main()
