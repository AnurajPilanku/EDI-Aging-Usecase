'''
Created by     : AnurajPilanku
Code utility   : remove rows based on a value in a column
Use case        : IPI2K Use case
'''

import openpyxl
import sys
import pandas as pd

#Groups
ng="US_NGG-APPL-Support"
es="US_EDI-Support"
ea="US_EDI-Analysts"
we="US_WebMethods-EAI-lvl3"

#path
excelPath=sys.argv[1]
mspath=sys.argv[2]
output=sys.argv[3]

#dframe
ngg=pd.read_excel(excelPath,sheet_name=0,index=False,engine='openpyxl')
ngg1=ngg[ngg['Assigned to Group'] != ng]
ediSupport=pd.read_excel(excelPath,sheet_name=1,index=False,engine='openpyxl')
ediSupport1=ediSupport[ediSupport['Assigned to Group'] != es]
ediAnalysts=pd.read_excel(excelPath,sheet_name=2,index=False,engine='openpyxl')
ediAnalysts1=ediAnalysts[ediAnalysts['Assigned to Group'] != ea]
webMethods=pd.read_excel(excelPath,sheet_name=3,index=False,engine='openpyxl')
webMethods1=webMethods[webMethods['Assigned to Group'] != we]

mssupport=pd.read_excel(mspath,index=False,engine='openpyxl')


#method1
'''
groups= []
ms_nonms=[]
for group in mssupport['group']:
    groups.append(group)

for value in ngg1['MS_NONMS']:
    if value in groups:
        ms_nonms.append("MS")
    else:
        ms_nonms.append("Fail")

ngg1["MS_NONMS"] = ms_nonms
'''

groups= []
for group in mssupport['group']:
    groups.append(group)

#ngg1['MS_NONMS']=ngg1['Assigned to Group'].apply(lambda x: 'MS' if x in mssupport['group'] else 'NON_MS')
ngg1['MS_NONMS'] = ['MS' if m in groups else 'Non_ms' for m in ngg1['Assigned to Group']]
ediSupport1['MS_NONMS'] = ['MS' if m in groups else 'Non_ms' for m in ediSupport1['Assigned to Group']]
ediAnalysts1['MS_NONMS'] = ['MS' if m in groups else 'Non_ms' for m in ediAnalysts1['Assigned to Group']]
webMethods1['MS_NONMS'] = ['MS' if m in groups else 'Non_ms' for m in webMethods1['Assigned to Group']]

#writer = pd.ExcelWriter(output, engine='xlsxwriter')

ngg1.to_excel(output.replace(".xlsx","_ngg.xlsx"),index=False)#, sheet_name=ng,index=False)
ediSupport1.to_excel(output.replace(".xlsx","_ediSupport.xlsx"),index=False)#(writer, sheet_name=es,index=False)
ediAnalysts1.to_excel(output.replace(".xlsx","_ediAnalysts.xlsx"),index=False)#(writer, sheet_name=ea,index=False)
webMethods1.to_excel(output.replace(".xlsx","_webmethods.xlsx"),index=False)#(writer, sheet_name=we,index=False)

#writer.save()
print("success")
#C:\Users\2040664\anuraj\EDI\ediData.xlsx
#US_NGG-APPL-Support
#US_EDI-Support
#US_EDI-Analysts
#US_WebMethods-EAI-lvl3