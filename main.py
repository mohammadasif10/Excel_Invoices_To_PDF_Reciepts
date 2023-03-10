import pandas as pd#can read csv AND excel
import glob
from fpdf import FPDF
from pathlib import Path#takes a filepath and can do things to the name
import openpyxl#you can just pip install this, inlcuded it here for refrence

filepaths = glob.glob("invoices/*.xlsx")#glob looks for patterns. * is a placeholder. it means anything can be here and since .xlsx is after it. it looping through a folder that has invoices/ at the beginning, anything in the middle, but .xlsx at the end and stores as a list. doesn't have to be .xlsx, can jsut be xlsx
print(filepaths)

for fp in filepaths: #the excel files in this case are called "dataframes"
  excelReader = pd.read_excel(fp, sheet_name="Sheet 1")#panda is reading the excel sheet that's  being looped over rn, have to give a sheet name bc excel can have diff sheets
  mypdf = FPDF(orientation="P", unit="mm", format="A4")
  mypdf.add_page()
  mypdf.set_font(family="Times", size=16, style="B")
  filename = Path(fp).stem#path from pathlib, takes the filepath and with stem, removes the invoices/ and the .xlsx, we are left w invoice number and date. just splitting would take too long bc splitting first at / and then xlsx. this does it one step, more useful on longer file names
  invNum = filename.split("-")[0]#split stores both halves in a list, we want the first part so we add the [0]
  date = filename.split("-")[1]
  mypdf.cell(w=50, h=8, txt=f"Invoice Number: {invNum}")
  mypdf.output(f"reciepts/{filename}.pdf") #since filename is saved each loop, this allows us to grab the filename of the file we are looping over and make a pdf with the same name. have to add .pdf at the end
  