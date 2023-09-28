from PyPDF2 import PdfReader
import openpyxl as xl
import pandas as pd

wb = xl.Workbook()
ws = wb.active


def create_string_within_range(lst, start_word, end_word):
    output_string = ""
    add_item = False
    
    for item in lst:
        if add_item and item != end_word:
            output_string += item + " "
        
        if item == start_word:
            add_item = True
        elif item == end_word:
            add_item = False
    
    return output_string.strip()

whole = []
ordered = []
final = []



studyNumber = ["Study Number"]
issuedDate = ["Date Issued"]
structureNumber = ["Structure Number"]
latitude = ["Latitude"]
longitude = ["Logitude"]
siteElevation = ["Site Elevation"]
aboveGroundCrane = ["AGH"]
# expirationDate = ["Expires"]

reader = PdfReader('ChaseMecostaSplit.pdf')
# pages = reader.pages[0]
# text = pages.extract_text()
# whole.append(text.split())
# print(whole)
for i in range(0,819):
    page = reader.pages[i]
    text = page.extract_text()
    pagei = text.split()
    if pagei[0] == "Mail":
        whole.append(pagei)
print(whole[0])

for k in whole: # get structure and crane number IDs
    for j in range(len(k)-1):
        if k[j]=="Study" and k[j+1] == "No.":
            studyNumber.append(k[j+2])


for j in whole: # get the issued Date
    for k in range(len(j)-1):
        if j[k]=="Date:":
            issuedDate.append(j[k+1])



for w in whole: # gets the structure/crane numbers
    strucNum= create_string_within_range(w, "Structure:", "Location:")
    structureNumber.append(strucNum)
    strucNum=""


for q in whole: # gets the latitudes
    for j in range(len(q)-1):
        if q[j] == "Latitude:":
            latitude.append(q[j+1][:-1])


for u in whole: # gets the longitudes
    for j in range(len(u)):
        if u[j] == "Longitude:":
            longitude.append(u[j+1][:-1])


for p in whole: # gets the site elevations
    for j in range(len(p)-1):
        if p[j] == "feet" and p[j+1] == "site" and p[j+2] == "elevation":
            siteElevation.append(p[j-1])

for a in whole: # above ground elevation
    for j in range(len(a)-1):
        if a[j] == "feet" and a[j+1] == "above" and a[j+1] == "ground":
            aboveGroundCrane.append(a[96]) #Fix this

# for i in whole:
#     index1 = whole.index(i)
#     for j in i:
#         index2 = i.index(j)
#         if j == "expires" and whole[index1][index2+1] == "on":
#             expirationDate.append(whole[index1][index2+2])



final.append(studyNumber)
final.append(issuedDate)
final.append(structureNumber)
final.append(latitude)
final.append(longitude)
final.append(siteElevation)
final.append(aboveGroundCrane)
# final.append(expirationDate)


final_transposed = list(map(list, zip(*final)))
df = pd.DataFrame(final_transposed)
df.to_excel("Cranes.xlsx")








    

















            
