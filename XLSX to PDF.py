#XLSX to PDF Created By MJ

# First download those libraries in your workspace with the command pip in your OS terminal
import openpyxl
import weasyprint

# Write the path of the filein th qouted space
workbook = openpyxl.load_workbook(" Name.xlsx ")

# Select the first worksheet
worksheet = workbook.worksheets[0]

# Get the printable area of the worksheet
print_area = worksheet.print_area

# Convert the print area into list of cell ranges
cell_ranges = print_area.split(",")

# Create a list of HTML tables, one for each cell range
tables = []
for cell_range in cell_ranges:
    table = "<table>"
    for row in worksheet[cell_range]:
        table += "<tr>"
        for cell in row:
            table += "<td>{}</td>".format(cell.value)
        table += "</tr>"
    table += "</table>"
    tables.append(table)

# Integrate the HTML tables into a single HTML string
html = "\n".join(tables)

# Convert the HTML to a PDF using the library WeasyPrint
pdf = weasyprint.HTML(string=html).write_pdf("Name.pdf")