# Importing neccesary libraries
import pandas as pd
from fpdf import FPDF

# Loading the CSV file
df = pd.read_csv("C:/Python Intership/Task2/SampleCSVFile.csv", encoding="ISO-8859-1")

# Basic Analysis of the file
total_records = len(df)
top_product = df["Product"].value_counts().idxmax()
top_customer = df["Name"].value_counts().idxmax()
top_category = df["Category"].value_counts().idxmax()
top_5_products = df["Product"].value_counts().head(5)

# Create PDF Report
pdf = FPDF()
pdf.set_auto_page_break(auto=True, margin=10)
pdf.add_page()
pdf.rect(x=5.0, y=5.0, w=200.0, h=287.0)  # Setting the page border for the PDF

# Setting font and size for the title of the report
pdf.set_font(family="Times", style="UB", size=18)
Title1 = pdf.cell(200, 10, txt="Product Sales Report", ln=True, align='C')
pdf.ln(10)

# Settign font and size for other text beside title 
pdf.set_font(family="Courier", size=12, style="B")
pdf.cell(200, 10, txt=f"Total Records: {total_records}", ln=True)
pdf.cell(200, 10, txt=f"Top Customer: {top_customer}", ln=True)
pdf.cell(200, 10, txt=f"Most Common Product: {top_product}", ln=True)
pdf.cell(200, 10, txt=f"Most Common Category: {top_category}", ln=True)


pdf.set_font(family="Times", style="UB", size=14)
pdf.ln(10)
pdf.cell(200, 10, txt="Top 5 Most Sold Products:", ln=True)


pdf.set_font(family="Courier", size=12, style="B")
for product, count in top_5_products.items():
    pdf.cell(200, 10, txt=f"- {product} ({count} times)", ln=True)

#  Adding a New Page for Full Sheet Table 
pdf.set_left_margin(2)
pdf.set_right_margin(2)
pdf.set_top_margin(2)
pdf.add_page()
pdf.set_font("Helvetica", "B", 14)
pdf.cell(200, 10, "Full Data Table", ln=True)
pdf.set_font("Helvetica", "", 10)
pdf.ln(5)

#  Table Header 
col_widths = [15, 105, 35, 45]  # Settign the width of columns in the table in Number, Product, Name, Category order
row_height = 10

# Printing the file data in the PDF table
for i, col in enumerate(df.columns):
    pdf.cell(col_widths[i], row_height, col, border=1)
    pdf.set_font(size=8)
pdf.ln(row_height)

#  Table Rows  
for i in range(len(df)):
    for j, item in enumerate(df.iloc[i]):
        pdf.cell(col_widths[j], row_height, str(item), border=1)
        pdf.set_font(size=8)
    pdf.ln(row_height)

# Output
pdf.output("Sales_Report.pdf")
print("Report generated successfully!")