import pdftables_api
import os

path = os.path.dirname(os.path.abspath(__file__))
pdffolder = os.path.join(path, "pdfs")
xfolder = os.path.join(path, "excels")
c = pdftables_api.Client('')

files = []
# r=root, d=directories, f = files
for r, d, f in os.walk(pdffolder):
    for file in f:
        if '.pdf' in file:
            files.append(os.path.join(r, file))

print(len(files))
for f in files:
    filename = f.split("/")[-1].split(".pdf")[0]
    output = os.path.join(xfolder, filename)
    if os.path.isfile(output+".xlsx") is not True:
        c.xlsx(f, output)
        print("convreted: ", output)
    else:
        print("exists")
