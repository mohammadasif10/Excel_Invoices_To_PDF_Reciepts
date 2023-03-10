import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")#glob looks for patterns. * is a placeholder. it means anything can be here and since .xlsx is after it. ot looks for anything that ends with xlsx and stores as a list. doesn't have to be .xlsx, can jsut be xlsx
print(filepaths)