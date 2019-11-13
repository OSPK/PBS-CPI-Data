import os        
import xlrd
import csv

path = os.path.dirname(os.path.abspath(__file__))

xfolder = os.path.join(path, "excels")
csvfolder = os.path.join(path, "csvs")

files = []
# r=root, d=directories, f = files
for r, d, f in os.walk(xfolder):
    for file in f:
        if '.xlsx' in file:
            files.append(os.path.join(r, file))

def csv_from_excel(inputf, outputf):
    wb = xlrd.open_workbook(inputf)
    sh = wb.sheet_by_name('Page 1')
    your_csv_file = open(output, 'w')
    wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

    for rownum in range(sh.nrows):
        wr.writerow(sh.row_values(rownum))

    your_csv_file.close()


for f in files:
    filename = f.split("/")[-1].split(".xlsx")[0]+".csv"
    output = os.path.join(csvfolder, filename)
    if os.path.isfile(output) is not True:
        print("convreting: ", f)
        csv_from_excel(f, output+".csv")
        print("convreted: ", output)
    else:
        print("exists")
# runs the csv_from_excel function:
