import fitz  # PyMuPDF
import pandas as pd


def pdf_to_excel(pdf_path, excel_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)

    # Create a list to hold PDF data
    data = []

    # Iterate through each page of the PDF
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        # Split the text into lines and add to the data list
        lines = text.split("\n")
        data.append(lines)

    # Create a DataFrame from the data
    df = pd.DataFrame(data).transpose()  # Transpose so that each line becomes a row

    # Save the DataFrame to an Excel file
    df.to_excel(excel_path, index=False, header=False)

    # Close the PDF
    pdf_document.close()
    print(f"Successfully converted {pdf_path} to {excel_path}")


# Example usage
pdf_path = "sample.pdf"
excel_path = "output.xlsx"
pdf_to_excel(pdf_path, excel_path)
