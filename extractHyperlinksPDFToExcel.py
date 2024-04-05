# import fitz  # PyMuPDF
# from openpyxl import Workbook
# from openpyxl.styles import Font, Alignment

# def extract_hyperlink_text(pdf_path):
#     doc = fitz.open(pdf_path)
#     hyperlinks_info = []

#     for page_num, page in enumerate(doc, start=1):
#         links = page.get_links()
#         for link in links:
#             if link['kind'] == fitz.LINK_URI:
#                 rect = fitz.Rect(link['from'])
#                 words = page.get_text("words")
#                 link_text = " ".join([w[4] for w in words if fitz.Rect(w[:4]).intersects(rect)])
#                 hyperlink_url = link['uri']
#                 hyperlinks_info.append((link_text.strip(), hyperlink_url, page_num))

#     doc.close()
#     return hyperlinks_info

# def save_hyperlinks_to_excel(hyperlinks_info, excel_path):
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Hyperlinks"
#     ws.append(["Hyperlink Text", "Hyperlink URL", "Page Number"])  # Adding the header
#     ws = wb.active  # Assuming 'wb' is your Workbook instance and 'ws' is the active worksheet
#     for cell in ws[1]:  # Accessing the first row directly with ws[1]
#         cell.font = Font(bold=True)
#         cell.alignment = Alignment(horizontal='center')

#     for row, info in enumerate(hyperlinks_info, start=2):  # Start from row 2 to account for the header
#         hyperlink_text, hyperlink_url, page_number = info
#         # Insert the hyperlink text and page number
#         ws[f'A{row}'] = hyperlink_text
#         ws[f'C{row}'] = page_number
#         ws[f'C{row}'].alignment = Alignment(horizontal='center')
#         # Create a clickable hyperlink for the URL
#         ws[f'B{row}'].hyperlink = hyperlink_url
#         ws[f'B{row}'].value = hyperlink_url  # You can customize this text
#         ws[f'B{row}'].style = "Hyperlink"  # Optional: Apply the hyperlink style

#     wb.save(excel_path)

# # Example usage
# pdf_path = "vishal.pdf"
# excel_path = "hyperlink_text.xlsx"
# hyperlinks_info = extract_hyperlink_text(pdf_path)
# save_hyperlinks_to_excel(hyperlinks_info, excel_path)


import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

def extract_hyperlink_text(pdf_path):
    doc = fitz.open(pdf_path)
    hyperlinks_info = []

    for page_num, page in enumerate(doc, start=1):
        links = page.get_links()
        for link in links:
            if link['kind'] == fitz.LINK_URI:
                rect = fitz.Rect(link['from'])
                words = page.get_text("words")
                link_text = " ".join([w[4] for w in words if fitz.Rect(w[:4]).intersects(rect)])
                hyperlink_url = link['uri']
                hyperlinks_info.append((link_text.strip(), hyperlink_url, page_num))

    doc.close()
    return hyperlinks_info

def save_hyperlinks_to_excel(hyperlinks_info, excel_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Hyperlinks"
    ws.append(["Page Number", "Hyperlink Text", "Hyperlink URL"])  # Adjusted the header order

    for cell in ws[1]:  # Formatting the header row
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    for row, info in enumerate(hyperlinks_info, start=2):
        hyperlink_text, hyperlink_url, page_number = info
        # Adjust the order here: Page number, Hyperlink text, and then Hyperlink URL
        ws[f'A{row}'] = page_number
        ws[f'A{row}'].alignment = Alignment(horizontal='center')  # Center-align the page number
        ws[f'B{row}'] = hyperlink_text
        # Create a clickable hyperlink for the URL
        ws[f'C{row}'].hyperlink = hyperlink_url
        ws[f'C{row}'].value = hyperlink_url  # You can customize this text
        ws[f'C{row}'].style = "Hyperlink"  # Optional: Apply the hyperlink style

    wb.save(excel_path)

# Example usage
pdf_path = "vishal.pdf"
excel_path = "hyperlink_text.xlsx"
hyperlinks_info = extract_hyperlink_text(pdf_path)
save_hyperlinks_to_excel(hyperlinks_info, excel_path)