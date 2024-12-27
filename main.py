from fpdf import FPDF
import pandas as pd
import os
import glob
from pathlib import Path

excel_files = glob.glob(os.path.join("invoices", "*.xlsx"))


for f in excel_files:
    
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()     
    filename= Path(f).stem
    invoice_nr= filename.split("-")[0]
    pdf.set_font("Times", size=16, style="B")
    pdf.cell(50, 10, txt= f"Invoice nr {invoice_nr}", ln=1, align="L")

    pdf.set_font("Times", size=16, style="B")
    pdf.cell(50, 10, txt= f"Date: {filename.split('-')[1]}", ln=1, align="L")    
    # create table
    df = pd.read_excel(f, sheet_name="Sheet 1")
    
    columns = df.columns
   
    pdf.set_font("Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(30, 10, txt= str(columns[0]), border=1)
    pdf.cell(60, 10, txt= str(columns[1]), border=1)
    pdf.cell(30, 10, txt= str(columns[2]),  border=1)
    pdf.cell(30, 10, txt= str(columns[3]),  border=1)
    pdf.cell(30, 10, txt= str(columns[4]),  border=1, ln=1)    
    
    for i, row in df.iterrows():
        
        pdf.set_font("Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(30, 10, txt= str(row["product_id"]), align="L", border=1)
        pdf.cell(60, 10, txt= str(row["product_name"]), align="L", border=1)
        pdf.cell(30, 10, txt= str(row["amount_purchased"]),  align="L", border=1)
        pdf.cell(30, 10, txt= str(row["price_per_unit"]),  align="L", border=1)
        pdf.cell(30, 10, txt= str(row["total_price"]),  align="L", border=1, ln=1)
       
   
    total_sum = df["total_price"].sum() 
    pdf.cell(30, 10, align="L", border=1)
    pdf.cell(60, 10, align="L", border=1)
    pdf.cell(30, 10, align="L", border=1)
    pdf.cell(30, 10, align="L", border=1)        
    pdf.cell(30, 10, txt= str(total_sum), border=1, ln=1)


    pdf.ln(4)
    pdf.set_font("Times", size=14, style="B")
    pdf.cell(30, 10, txt= f"Total price due: {total_sum} dollors", ln=1, align="L")
    pdf.cell(25, 10, txt= "Python How",  align="L")
   
    pdf.image("pythonhow.png", w=10)

    if not os.path.exists("PDFs"):
        os.mkdir("PDFs")   
    pdf.output(f"PDFs/{filename}.pdf")     
    



