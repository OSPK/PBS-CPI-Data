import camelot
import os
path = os.path.dirname(os.path.abspath(__file__))
tables = camelot.read_pdf(path+'/pdfs/2016-september.pdf')
tables.export(path+'/2016-september.csv', f='csv', compress=True)