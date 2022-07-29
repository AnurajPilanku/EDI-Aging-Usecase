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
rowpath=sys.argv[2]

#dframe
ean=pd.read_excel(excelPath,sheet_name=0,index=False,engine='openpyxl')

#MS_NONMS
Msean=ean[ean['MS_NONMS'] == "MS"]
NonMsean=ean[ean['MS_NONMS'] == "Non_ms"]
#MS Status
eanMSOpen=Msean[Msean['Status'] == "Open"]
eanMSPending_Client=Msean[Msean['Status'] == "Pending Client"]
eanMSPending_Emergency_Client=Msean[Msean['Status'] == "Pending Emergency Client"]
eanMSPending_Vendor=Msean[Msean['Status'] == "Pending Vendor"]
eanMSWork_In_Progress=Msean[Msean['Status'] == "Work In Progress"]
#(eanMSWork_In_Progress)
#Non_MS Status
eanNonMSOpen=NonMsean[NonMsean['Status'] == "Open"]
eanNonMSPending_Client=NonMsean[NonMsean['Status'] == "Pending Client"]
eanNonMSPending_Emergency_Client=NonMsean[NonMsean['Status'] == "Pending Emergency Client"]
eanNonMSPending_Vendor=NonMsean[NonMsean['Status'] == "Pending Vendor"]
eanNonMSWork_In_Progress=NonMsean[NonMsean['Status'] == "Work In Progress"]

#MS_OPEN_PRIORITY**
eanMSOpenCritical=eanMSOpen[eanMSOpen['Priority'] == 1]
eanMSOpenHigh=eanMSOpen[eanMSOpen['Priority'] == 2]
eanMSOpenAverage=eanMSOpen[eanMSOpen['Priority'] == 3]
eanMSOpenLow=eanMSOpen[eanMSOpen['Priority'] == 4]

##MS_Pending_Client_PRIORITY
eanMSPending_ClientCritical=eanMSPending_Client[eanMSPending_Client['Priority'] == 1]
eanMSPending_ClientHigh=eanMSPending_Client[eanMSPending_Client['Priority'] == 2]
eanMSPending_ClientAverage=eanMSPending_Client[eanMSPending_Client['Priority'] == 3]
eanMSPending_ClientLow=eanMSPending_Client[eanMSPending_Client['Priority'] == 4]

#MS_Pending_Emergency_Client
eanMSPending_Emergency_ClientCritical=eanMSPending_Emergency_Client[eanMSPending_Emergency_Client['Priority'] == 1]
eanMSPending_Emergency_ClientHigh=eanMSPending_Emergency_Client[eanMSPending_Emergency_Client['Priority'] == 2]
eanMSPending_Emergency_ClientAverage=eanMSPending_Emergency_Client[eanMSPending_Emergency_Client['Priority'] == 3]
eanMSPending_Emergency_ClientLow=eanMSPending_Emergency_Client[eanMSPending_Emergency_Client['Priority'] == 4]

#MS_Pending_Vendor
eanMSPending_VendorCritical=eanMSPending_Vendor[eanMSPending_Vendor['Priority'] == 1]
eanMSPending_VendorHigh=eanMSPending_Vendor[eanMSPending_Vendor['Priority'] == 2]
eanMSPending_VendorAverage=eanMSPending_Vendor[eanMSPending_Vendor['Priority'] == 3]
eanMSPending_VendorLow=eanMSPending_Vendor[eanMSPending_Vendor['Priority'] == 4]

#MS_Work_In_Progress
eanMSWork_In_ProgressCritical=eanMSWork_In_Progress[eanMSWork_In_Progress['Priority'] == 1]
eanMSWork_In_ProgressHigh=eanMSWork_In_Progress[eanMSWork_In_Progress['Priority'] == 2]
eanMSWork_In_ProgressAverage=eanMSWork_In_Progress[eanMSWork_In_Progress['Priority'] == 3]
eanMSWork_In_ProgressLow=eanMSWork_In_Progress[eanMSWork_In_Progress['Priority'] == 4]
#(eanMSWork_In_ProgressHigh)


#NonMS_OPEN_PRIORITY
eanNonMSOpenCritical=eanNonMSOpen[eanNonMSOpen['Priority'] == 1]
eanNonMSOpenHigh=eanNonMSOpen[eanNonMSOpen['Priority'] == 2]
eanNonMSOpenAverage=eanNonMSOpen[eanNonMSOpen['Priority'] == 3]
eanNonMSOpenLow=eanNonMSOpen[eanNonMSOpen['Priority'] == 4]

##NonMS_Pending_Client_PRIORITY
eanNonMSPending_ClientCritical=eanNonMSPending_Client[eanNonMSPending_Client['Priority'] == 1]
eanNonMSPending_ClientHigh=eanNonMSPending_Client[eanNonMSPending_Client['Priority'] == 2]
eanNonMSPending_ClientAverage=eanNonMSPending_Client[eanNonMSPending_Client['Priority'] == 3]
eanNonMSPending_ClientLow=eanNonMSPending_Client[eanNonMSPending_Client['Priority'] == 4]

#NonMS_Pending_Emergency_Client
eanNonMSPending_Emergency_ClientCritical=eanNonMSPending_Emergency_Client[eanNonMSPending_Emergency_Client['Priority'] == 1]
eanNonMSPending_Emergency_ClientHigh=eanNonMSPending_Emergency_Client[eanNonMSPending_Emergency_Client['Priority'] == 2]
eanNonMSPending_Emergency_ClientAverage=eanNonMSPending_Emergency_Client[eanNonMSPending_Emergency_Client['Priority'] == 3]
eanNonMSPending_Emergency_ClientLow=eanNonMSPending_Emergency_Client[eanNonMSPending_Emergency_Client['Priority'] == 4]

#NonMS_Pending_Vendor
eanNonMSPending_VendorCritical=eanNonMSPending_Vendor[eanNonMSPending_Vendor['Priority'] == 1]
eanNonMSPending_VendorHigh=eanNonMSPending_Vendor[eanNonMSPending_Vendor['Priority'] == 2]
eanNonMSPending_VendorAverage=eanNonMSPending_Vendor[eanNonMSPending_Vendor['Priority'] == 3]
eanNonMSPending_VendorLow=eanNonMSPending_Vendor[eanNonMSPending_Vendor['Priority'] == 4]

#NonMS_Work_In_Progress
eanNonMSWork_In_ProgressCritical=eanNonMSWork_In_Progress[eanNonMSWork_In_Progress['Priority'] == 1]
eanNonMSWork_In_ProgressHigh=eanNonMSWork_In_Progress[eanNonMSWork_In_Progress['Priority'] == 2]
eanNonMSWork_In_ProgressAverage=eanNonMSWork_In_Progress[eanNonMSWork_In_Progress['Priority'] == 3]
eanNonMSWork_In_ProgressLow=eanNonMSWork_In_Progress[eanNonMSWork_In_Progress['Priority'] == 4]

#eanMSOpenCriticalAging
eanMSOpenCriticalLessThan10=eanMSOpenCritical[eanMSOpenCritical['Ageing days'] <=10]
eanMSOpenCriticalGreaterThan10=eanMSOpenCritical[eanMSOpenCritical['Ageing days'] > 10]
eanMSOpenCriticalGreaterThan50=eanMSOpenCritical[eanMSOpenCritical['Ageing days'] > 50]
eanMSOpenCriticalGreaterThan100=eanMSOpenCritical[eanMSOpenCritical['Ageing days'] > 100]

#eanMSOpenHighAging
eanMSOpenHighLessThan10=eanMSOpenHigh[eanMSOpenHigh['Ageing days'] <=10]
eanMSOpenHighGreaterThan10=eanMSOpenHigh[eanMSOpenHigh['Ageing days'] > 10]
eanMSOpenHighGreaterThan50=eanMSOpenHigh[eanMSOpenHigh['Ageing days'] > 50]
eanMSOpenHighGreaterThan100=eanMSOpenHigh[eanMSOpenHigh['Ageing days'] > 100]

#eanMSOpenAverageAging
eanMSOpenAverageLessThan10=eanMSOpenAverage[eanMSOpenAverage['Ageing days'] <=10]
eanMSOpenAverageGreaterThan10=eanMSOpenAverage[eanMSOpenAverage['Ageing days'] > 10]
eanMSOpenAverageGreaterThan50=eanMSOpenAverage[eanMSOpenAverage['Ageing days'] > 50]
eanMSOpenAverageGreaterThan100=eanMSOpenAverage[eanMSOpenAverage['Ageing days'] > 100]

#eanMSOpenLowAging
eanMSOpenLowLessThan10=eanMSOpenLow[eanMSOpenLow['Ageing days'] <=10]
eanMSOpenLowGreaterThan10=eanMSOpenLow[eanMSOpenLow['Ageing days'] > 10]
eanMSOpenLowGreaterThan50=eanMSOpenLow[eanMSOpenLow['Ageing days'] > 50]
eanMSOpenLowGreaterThan100=eanMSOpenLow[eanMSOpenLow['Ageing days'] > 100]


#eanMSPending_ClientCriticalAging
eanMSPending_ClientCriticalLessThan10=eanMSPending_ClientCritical[eanMSPending_ClientCritical['Ageing days'] <=10]
eanMSPending_ClientCriticalGreaterThan10=eanMSPending_ClientCritical[eanMSPending_ClientCritical['Ageing days'] > 10]
eanMSPending_ClientCriticalGreaterThan50=eanMSPending_ClientCritical[eanMSPending_ClientCritical['Ageing days'] > 50]
eanMSPending_ClientCriticalGreaterThan100=eanMSPending_ClientCritical[eanMSPending_ClientCritical['Ageing days'] > 100]

#eanMSPending_ClientHighAging
eanMSPending_ClientHighLessThan10=eanMSPending_ClientHigh[eanMSPending_ClientHigh['Ageing days'] <=10]
eanMSPending_ClientHighGreaterThan10=eanMSPending_ClientHigh[eanMSPending_ClientHigh['Ageing days'] > 10]
eanMSPending_ClientHighGreaterThan50=eanMSPending_ClientHigh[eanMSPending_ClientHigh['Ageing days'] > 50]
eanMSPending_ClientHighGreaterThan100=eanMSPending_ClientHigh[eanMSPending_ClientHigh['Ageing days'] > 100]

#eanMSPending_ClientAverageAging
eanMSPending_ClientAverageLessThan10=eanMSPending_ClientAverage[eanMSPending_ClientAverage['Ageing days'] <=10]
eanMSPending_ClientAverageGreaterThan10=eanMSPending_ClientAverage[eanMSPending_ClientAverage['Ageing days'] > 10]
eanMSPending_ClientAverageGreaterThan50=eanMSPending_ClientAverage[eanMSPending_ClientAverage['Ageing days'] > 50]
eanMSPending_ClientAverageGreaterThan100=eanMSPending_ClientAverage[eanMSPending_ClientAverage['Ageing days'] > 100]

#eanMSPending_ClientLowAging
eanMSPending_ClientLowLessThan10=eanMSPending_ClientLow[eanMSPending_ClientLow['Ageing days'] <=10]
eanMSPending_ClientLowGreaterThan10=eanMSPending_ClientLow[eanMSPending_ClientLow['Ageing days'] > 10]
eanMSPending_ClientLowGreaterThan50=eanMSPending_ClientLow[eanMSPending_ClientLow['Ageing days'] > 50]
eanMSPending_ClientLowGreaterThan100=eanMSPending_ClientLow[eanMSPending_ClientLow['Ageing days'] > 100]


#eanMSPending_Emergency_ClientCriticalAging
eanMSPending_Emergency_ClientCriticalLessThan10=eanMSPending_Emergency_ClientCritical[eanMSPending_Emergency_ClientCritical['Ageing days'] <=10]
eanMSPending_Emergency_ClientCriticalGreaterThan10=eanMSPending_Emergency_ClientCritical[eanMSPending_Emergency_ClientCritical['Ageing days'] > 10]
eanMSPending_Emergency_ClientCriticalGreaterThan50=eanMSPending_Emergency_ClientCritical[eanMSPending_Emergency_ClientCritical['Ageing days'] > 50]
eanMSPending_Emergency_ClientCriticalGreaterThan100=eanMSPending_Emergency_ClientCritical[eanMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#eanMSPending_Emergency_ClientHighAging
eanMSPending_Emergency_ClientHighLessThan10=eanMSPending_Emergency_ClientHigh[eanMSPending_Emergency_ClientHigh['Ageing days'] <=10]
eanMSPending_Emergency_ClientHighGreaterThan10=eanMSPending_Emergency_ClientHigh[eanMSPending_Emergency_ClientHigh['Ageing days'] > 10]
eanMSPending_Emergency_ClientHighGreaterThan50=eanMSPending_Emergency_ClientHigh[eanMSPending_Emergency_ClientHigh['Ageing days'] > 50]
eanMSPending_Emergency_ClientHighGreaterThan100=eanMSPending_Emergency_ClientHigh[eanMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#eanMSPending_Emergency_ClientAverageAging
eanMSPending_Emergency_ClientAverageLessThan10=eanMSPending_Emergency_ClientAverage[eanMSPending_Emergency_ClientAverage['Ageing days'] <=10]
eanMSPending_Emergency_ClientAverageGreaterThan10=eanMSPending_Emergency_ClientAverage[eanMSPending_Emergency_ClientAverage['Ageing days'] > 10]
eanMSPending_Emergency_ClientAverageGreaterThan50=eanMSPending_Emergency_ClientAverage[eanMSPending_Emergency_ClientAverage['Ageing days'] > 50]
eanMSPending_Emergency_ClientAverageGreaterThan100=eanMSPending_Emergency_ClientAverage[eanMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#eanMSPending_Emergency_ClientLowAging
eanMSPending_Emergency_ClientLowLessThan10=eanMSPending_Emergency_ClientLow[eanMSPending_Emergency_ClientLow['Ageing days'] <=10]
eanMSPending_Emergency_ClientLowGreaterThan10=eanMSPending_Emergency_ClientLow[eanMSPending_Emergency_ClientLow['Ageing days'] > 10]
eanMSPending_Emergency_ClientLowGreaterThan50=eanMSPending_Emergency_ClientLow[eanMSPending_Emergency_ClientLow['Ageing days'] > 50]
eanMSPending_Emergency_ClientLowGreaterThan100=eanMSPending_Emergency_ClientLow[eanMSPending_Emergency_ClientLow['Ageing days'] > 100]


#eanMSPending_VendorCriticalAging
eanMSPending_VendorCriticalLessThan10=eanMSPending_VendorCritical[eanMSPending_VendorCritical['Ageing days'] <=10]
eanMSPending_VendorCriticalGreaterThan10=eanMSPending_VendorCritical[eanMSPending_VendorCritical['Ageing days'] > 10]
eanMSPending_VendorCriticalGreaterThan50=eanMSPending_VendorCritical[eanMSPending_VendorCritical['Ageing days'] > 50]
eanMSPending_VendorCriticalGreaterThan100=eanMSPending_VendorCritical[eanMSPending_VendorCritical['Ageing days'] > 100]

#eanMSPending_VendorHighAging
eanMSPending_VendorHighLessThan10=eanMSPending_VendorHigh[eanMSPending_VendorHigh['Ageing days'] <=10]
eanMSPending_VendorHighGreaterThan10=eanMSPending_VendorHigh[eanMSPending_VendorHigh['Ageing days'] > 10]
eanMSPending_VendorHighGreaterThan50=eanMSPending_VendorHigh[eanMSPending_VendorHigh['Ageing days'] > 50]
eanMSPending_VendorHighGreaterThan100=eanMSPending_VendorHigh[eanMSPending_VendorHigh['Ageing days'] > 100]

#eanMSPending_VendorAverageAging
eanMSPending_VendorAverageLessThan10=eanMSPending_VendorAverage[eanMSPending_VendorAverage['Ageing days'] <=10]
eanMSPending_VendorAverageGreaterThan10=eanMSPending_VendorAverage[eanMSPending_VendorAverage['Ageing days'] > 10]
eanMSPending_VendorAverageGreaterThan50=eanMSPending_VendorAverage[eanMSPending_VendorAverage['Ageing days'] > 50]
eanMSPending_VendorAverageGreaterThan100=eanMSPending_VendorAverage[eanMSPending_VendorAverage['Ageing days'] > 100]

#eanMSPending_VendorLowAging
eanMSPending_VendorLowLessThan10=eanMSPending_VendorLow[eanMSPending_VendorLow['Ageing days'] <=10]
eanMSPending_VendorLowGreaterThan10=eanMSPending_VendorLow[eanMSPending_VendorLow['Ageing days'] > 10]
eanMSPending_VendorLowGreaterThan50=eanMSPending_VendorLow[eanMSPending_VendorLow['Ageing days'] > 50]
eanMSPending_VendorLowGreaterThan100=eanMSPending_VendorLow[eanMSPending_VendorLow['Ageing days'] > 100]


#eanMSWork_In_ProgressCriticalAging
eanMSWork_In_ProgressCriticalLessThan10=eanMSWork_In_ProgressCritical[eanMSWork_In_ProgressCritical['Ageing days'] <=10]
eanMSWork_In_ProgressCriticalGreaterThan10=eanMSWork_In_ProgressCritical[eanMSWork_In_ProgressCritical['Ageing days'] > 10]
eanMSWork_In_ProgressCriticalGreaterThan50=eanMSWork_In_ProgressCritical[eanMSWork_In_ProgressCritical['Ageing days'] > 50]
eanMSWork_In_ProgressCriticalGreaterThan100=eanMSWork_In_ProgressCritical[eanMSWork_In_ProgressCritical['Ageing days'] > 100]

#eanMSWork_In_ProgressHighAging
eanMSWork_In_ProgressHighLessThan10=eanMSWork_In_ProgressHigh[eanMSWork_In_ProgressHigh['Ageing days'] <=10]
eanMSWork_In_ProgressHighGreaterThan10=eanMSWork_In_ProgressHigh[eanMSWork_In_ProgressHigh['Ageing days'] > 10]
eanMSWork_In_ProgressHighGreaterThan50=eanMSWork_In_ProgressHigh[eanMSWork_In_ProgressHigh['Ageing days'] > 50]
eanMSWork_In_ProgressHighGreaterThan100=eanMSWork_In_ProgressHigh[eanMSWork_In_ProgressHigh['Ageing days'] > 100]
#(eanMSWork_In_ProgressHighLessThan10)
#(eanMSWork_In_ProgressHighLessThan10.shape[0])
#eanMSWork_In_ProgressAverageAging
eanMSWork_In_ProgressAverageLessThan10=eanMSWork_In_ProgressAverage[eanMSWork_In_ProgressAverage['Ageing days'] <=10]
eanMSWork_In_ProgressAverageGreaterThan10=eanMSWork_In_ProgressAverage[eanMSWork_In_ProgressAverage['Ageing days'] > 10]
eanMSWork_In_ProgressAverageGreaterThan50=eanMSWork_In_ProgressAverage[eanMSWork_In_ProgressAverage['Ageing days'] > 50]
eanMSWork_In_ProgressAverageGreaterThan100=eanMSWork_In_ProgressAverage[eanMSWork_In_ProgressAverage['Ageing days'] > 100]

#eanMSWork_In_ProgressLowAging
eanMSWork_In_ProgressLowLessThan10=eanMSWork_In_ProgressLow[eanMSWork_In_ProgressLow['Ageing days'] <=10]
eanMSWork_In_ProgressLowGreaterThan10=eanMSWork_In_ProgressLow[eanMSWork_In_ProgressLow['Ageing days'] > 10]
eanMSWork_In_ProgressLowGreaterThan50=eanMSWork_In_ProgressLow[eanMSWork_In_ProgressLow['Ageing days'] > 50]
eanMSWork_In_ProgressLowGreaterThan100=eanMSWork_In_ProgressLow[eanMSWork_In_ProgressLow['Ageing days'] > 100]



#eanNonMSOpenCriticalAging
eanNonMSOpenCriticalLessThan10=eanNonMSOpenCritical[eanNonMSOpenCritical['Ageing days'] <=10]
eanNonMSOpenCriticalGreaterThan10=eanNonMSOpenCritical[eanNonMSOpenCritical['Ageing days'] > 10]
eanNonMSOpenCriticalGreaterThan50=eanNonMSOpenCritical[eanNonMSOpenCritical['Ageing days'] > 50]
eanNonMSOpenCriticalGreaterThan100=eanNonMSOpenCritical[eanNonMSOpenCritical['Ageing days'] > 100]

#eanNonMSOpenHighAging
eanNonMSOpenHighLessThan10=eanNonMSOpenHigh[eanNonMSOpenHigh['Ageing days'] <=10]
eanNonMSOpenHighGreaterThan10=eanNonMSOpenHigh[eanNonMSOpenHigh['Ageing days'] > 10]
eanNonMSOpenHighGreaterThan50=eanNonMSOpenHigh[eanNonMSOpenHigh['Ageing days'] > 50]
eanNonMSOpenHighGreaterThan100=eanNonMSOpenHigh[eanNonMSOpenHigh['Ageing days'] > 100]

#eanNonMSOpenAverageAging
eanNonMSOpenAverageLessThan10=eanNonMSOpenAverage[eanNonMSOpenAverage['Ageing days'] <=10]
eanNonMSOpenAverageGreaterThan10=eanNonMSOpenAverage[eanNonMSOpenAverage['Ageing days'] > 10]
eanNonMSOpenAverageGreaterThan50=eanNonMSOpenAverage[eanNonMSOpenAverage['Ageing days'] > 50]
eanNonMSOpenAverageGreaterThan100=eanNonMSOpenAverage[eanNonMSOpenAverage['Ageing days'] > 100]

#eanNonMSOpenLowAging
eanNonMSOpenLowLessThan10=eanNonMSOpenLow[eanNonMSOpenLow['Ageing days'] <=10]
eanNonMSOpenLowGreaterThan10=eanNonMSOpenLow[eanNonMSOpenLow['Ageing days'] > 10]
eanNonMSOpenLowGreaterThan50=eanNonMSOpenLow[eanNonMSOpenLow['Ageing days'] > 50]
eanNonMSOpenLowGreaterThan100=eanNonMSOpenLow[eanNonMSOpenLow['Ageing days'] > 100]


#eanNonMSPending_ClientCriticalAging
eanNonMSPending_ClientCriticalLessThan10=eanNonMSPending_ClientCritical[eanNonMSPending_ClientCritical['Ageing days'] <=10]
eanNonMSPending_ClientCriticalGreaterThan10=eanNonMSPending_ClientCritical[eanNonMSPending_ClientCritical['Ageing days'] > 10]
eanNonMSPending_ClientCriticalGreaterThan50=eanNonMSPending_ClientCritical[eanNonMSPending_ClientCritical['Ageing days'] > 50]
eanNonMSPending_ClientCriticalGreaterThan100=eanNonMSPending_ClientCritical[eanNonMSPending_ClientCritical['Ageing days'] > 100]

#eanNonMSPending_ClientHighAging
eanNonMSPending_ClientHighLessThan10=eanNonMSPending_ClientHigh[eanNonMSPending_ClientHigh['Ageing days'] <=10]
eanNonMSPending_ClientHighGreaterThan10=eanNonMSPending_ClientHigh[eanNonMSPending_ClientHigh['Ageing days'] > 10]
eanNonMSPending_ClientHighGreaterThan50=eanNonMSPending_ClientHigh[eanNonMSPending_ClientHigh['Ageing days'] > 50]
eanNonMSPending_ClientHighGreaterThan100=eanNonMSPending_ClientHigh[eanNonMSPending_ClientHigh['Ageing days'] > 100]

#eanNonMSPending_ClientAverageAging
eanNonMSPending_ClientAverageLessThan10=eanNonMSPending_ClientAverage[eanNonMSPending_ClientAverage['Ageing days'] <=10]
eanNonMSPending_ClientAverageGreaterThan10=eanNonMSPending_ClientAverage[eanNonMSPending_ClientAverage['Ageing days'] > 10]
eanNonMSPending_ClientAverageGreaterThan50=eanNonMSPending_ClientAverage[eanNonMSPending_ClientAverage['Ageing days'] > 50]
eanNonMSPending_ClientAverageGreaterThan100=eanNonMSPending_ClientAverage[eanNonMSPending_ClientAverage['Ageing days'] > 100]

#eanNonMSPending_ClientLowAging
eanNonMSPending_ClientLowLessThan10=eanNonMSPending_ClientLow[eanNonMSPending_ClientLow['Ageing days'] <=10]
eanNonMSPending_ClientLowGreaterThan10=eanNonMSPending_ClientLow[eanNonMSPending_ClientLow['Ageing days'] > 10]
eanNonMSPending_ClientLowGreaterThan50=eanNonMSPending_ClientLow[eanNonMSPending_ClientLow['Ageing days'] > 50]
eanNonMSPending_ClientLowGreaterThan100=eanNonMSPending_ClientLow[eanNonMSPending_ClientLow['Ageing days'] > 100]


#eanNonMSPending_Emergency_ClientCriticalAging
eanNonMSPending_Emergency_ClientCriticalLessThan10=eanNonMSPending_Emergency_ClientCritical[eanNonMSPending_Emergency_ClientCritical['Ageing days'] <=10]
eanNonMSPending_Emergency_ClientCriticalGreaterThan10=eanNonMSPending_Emergency_ClientCritical[eanNonMSPending_Emergency_ClientCritical['Ageing days'] > 10]
eanNonMSPending_Emergency_ClientCriticalGreaterThan50=eanNonMSPending_Emergency_ClientCritical[eanNonMSPending_Emergency_ClientCritical['Ageing days'] > 50]
eanNonMSPending_Emergency_ClientCriticalGreaterThan100=eanNonMSPending_Emergency_ClientCritical[eanNonMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#eanNonMSPending_Emergency_ClientHighAging
eanNonMSPending_Emergency_ClientHighLessThan10=eanNonMSPending_Emergency_ClientHigh[eanNonMSPending_Emergency_ClientHigh['Ageing days'] <=10]
eanNonMSPending_Emergency_ClientHighGreaterThan10=eanNonMSPending_Emergency_ClientHigh[eanNonMSPending_Emergency_ClientHigh['Ageing days'] > 10]
eanNonMSPending_Emergency_ClientHighGreaterThan50=eanNonMSPending_Emergency_ClientHigh[eanNonMSPending_Emergency_ClientHigh['Ageing days'] > 50]
eanNonMSPending_Emergency_ClientHighGreaterThan100=eanNonMSPending_Emergency_ClientHigh[eanNonMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#eanNonMSPending_Emergency_ClientAverageAging
eanNonMSPending_Emergency_ClientAverageLessThan10=eanNonMSPending_Emergency_ClientAverage[eanNonMSPending_Emergency_ClientAverage['Ageing days'] <=10]
eanNonMSPending_Emergency_ClientAverageGreaterThan10=eanNonMSPending_Emergency_ClientAverage[eanNonMSPending_Emergency_ClientAverage['Ageing days'] > 10]
eanNonMSPending_Emergency_ClientAverageGreaterThan50=eanNonMSPending_Emergency_ClientAverage[eanNonMSPending_Emergency_ClientAverage['Ageing days'] > 50]
eanNonMSPending_Emergency_ClientAverageGreaterThan100=eanNonMSPending_Emergency_ClientAverage[eanNonMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#eanNonMSPending_Emergency_ClientLowAging
eanNonMSPending_Emergency_ClientLowLessThan10=eanNonMSPending_Emergency_ClientLow[eanNonMSPending_Emergency_ClientLow['Ageing days'] <=10]
eanNonMSPending_Emergency_ClientLowGreaterThan10=eanNonMSPending_Emergency_ClientLow[eanNonMSPending_Emergency_ClientLow['Ageing days'] > 10]
eanNonMSPending_Emergency_ClientLowGreaterThan50=eanNonMSPending_Emergency_ClientLow[eanNonMSPending_Emergency_ClientLow['Ageing days'] > 50]
eanNonMSPending_Emergency_ClientLowGreaterThan100=eanNonMSPending_Emergency_ClientLow[eanNonMSPending_Emergency_ClientLow['Ageing days'] > 100]


#eanNonMSPending_VendorCriticalAging
eanNonMSPending_VendorCriticalLessThan10=eanNonMSPending_VendorCritical[eanNonMSPending_VendorCritical['Ageing days'] <=10]
eanNonMSPending_VendorCriticalGreaterThan10=eanNonMSPending_VendorCritical[eanNonMSPending_VendorCritical['Ageing days'] > 10]
eanNonMSPending_VendorCriticalGreaterThan50=eanNonMSPending_VendorCritical[eanNonMSPending_VendorCritical['Ageing days'] > 50]
eanNonMSPending_VendorCriticalGreaterThan100=eanNonMSPending_VendorCritical[eanNonMSPending_VendorCritical['Ageing days'] > 100]

#eanNonMSPending_VendorHighAging
eanNonMSPending_VendorHighLessThan10=eanNonMSPending_VendorHigh[eanNonMSPending_VendorHigh['Ageing days'] <=10]
eanNonMSPending_VendorHighGreaterThan10=eanNonMSPending_VendorHigh[eanNonMSPending_VendorHigh['Ageing days'] > 10]
eanNonMSPending_VendorHighGreaterThan50=eanNonMSPending_VendorHigh[eanNonMSPending_VendorHigh['Ageing days'] > 50]
eanNonMSPending_VendorHighGreaterThan100=eanNonMSPending_VendorHigh[eanNonMSPending_VendorHigh['Ageing days'] > 100]

#eanNonMSPending_VendorAverageAging
eanNonMSPending_VendorAverageLessThan10=eanNonMSPending_VendorAverage[eanNonMSPending_VendorAverage['Ageing days'] <=10]
eanNonMSPending_VendorAverageGreaterThan10=eanNonMSPending_VendorAverage[eanNonMSPending_VendorAverage['Ageing days'] > 10]
eanNonMSPending_VendorAverageGreaterThan50=eanNonMSPending_VendorAverage[eanNonMSPending_VendorAverage['Ageing days'] > 50]
eanNonMSPending_VendorAverageGreaterThan100=eanNonMSPending_VendorAverage[eanNonMSPending_VendorAverage['Ageing days'] > 100]

#eanNonMSPending_VendorLowAging
eanNonMSPending_VendorLowLessThan10=eanNonMSPending_VendorLow[eanNonMSPending_VendorLow['Ageing days'] <=10]
eanNonMSPending_VendorLowGreaterThan10=eanNonMSPending_VendorLow[eanNonMSPending_VendorLow['Ageing days'] > 10]
eanNonMSPending_VendorLowGreaterThan50=eanNonMSPending_VendorLow[eanNonMSPending_VendorLow['Ageing days'] > 50]
eanNonMSPending_VendorLowGreaterThan100=eanNonMSPending_VendorLow[eanNonMSPending_VendorLow['Ageing days'] > 100]


#eanNonMSWork_In_ProgressCriticalAging
eanNonMSWork_In_ProgressCriticalLessThan10=eanNonMSWork_In_ProgressCritical[eanNonMSWork_In_ProgressCritical['Ageing days'] <=10]
eanNonMSWork_In_ProgressCriticalGreaterThan10=eanNonMSWork_In_ProgressCritical[eanNonMSWork_In_ProgressCritical['Ageing days'] > 10]
eanNonMSWork_In_ProgressCriticalGreaterThan50=eanNonMSWork_In_ProgressCritical[eanNonMSWork_In_ProgressCritical['Ageing days'] > 50]
eanNonMSWork_In_ProgressCriticalGreaterThan100=eanNonMSWork_In_ProgressCritical[eanNonMSWork_In_ProgressCritical['Ageing days'] > 100]

#eanNonMSWork_In_ProgressHighAging
eanNonMSWork_In_ProgressHighLessThan10=eanNonMSWork_In_ProgressHigh[eanNonMSWork_In_ProgressHigh['Ageing days'] <=10]
eanNonMSWork_In_ProgressHighGreaterThan10=eanNonMSWork_In_ProgressHigh[eanNonMSWork_In_ProgressHigh['Ageing days'] > 10]
eanNonMSWork_In_ProgressHighGreaterThan50=eanNonMSWork_In_ProgressHigh[eanNonMSWork_In_ProgressHigh['Ageing days'] > 50]
eanNonMSWork_In_ProgressHighGreaterThan100=eanNonMSWork_In_ProgressHigh[eanNonMSWork_In_ProgressHigh['Ageing days'] > 100]

#eanNonMSWork_In_ProgressAverageAging
eanNonMSWork_In_ProgressAverageLessThan10=eanNonMSWork_In_ProgressAverage[eanNonMSWork_In_ProgressAverage['Ageing days'] <=10]
eanNonMSWork_In_ProgressAverageGreaterThan10=eanNonMSWork_In_ProgressAverage[eanNonMSWork_In_ProgressAverage['Ageing days'] > 10]
eanNonMSWork_In_ProgressAverageGreaterThan50=eanNonMSWork_In_ProgressAverage[eanNonMSWork_In_ProgressAverage['Ageing days'] > 50]
eanNonMSWork_In_ProgressAverageGreaterThan100=eanNonMSWork_In_ProgressAverage[eanNonMSWork_In_ProgressAverage['Ageing days'] > 100]

#eanNonMSWork_In_ProgressLowAging
eanNonMSWork_In_ProgressLowLessThan10=eanNonMSWork_In_ProgressLow[eanNonMSWork_In_ProgressLow['Ageing days'] <=10]
eanNonMSWork_In_ProgressLowGreaterThan10=eanNonMSWork_In_ProgressLow[eanNonMSWork_In_ProgressLow['Ageing days'] > 10]
eanNonMSWork_In_ProgressLowGreaterThan50=eanNonMSWork_In_ProgressLow[eanNonMSWork_In_ProgressLow['Ageing days'] > 50]
eanNonMSWork_In_ProgressLowGreaterThan100=eanNonMSWork_In_ProgressLow[eanNonMSWork_In_ProgressLow['Ageing days'] > 100]


rowcount={'Data':['eanMSOpenCriticalLessThan10','eanMSOpenCriticalGreaterThan10','eanMSOpenCriticalGreaterThan50','eanMSOpenCriticalGreaterThan100','eanMSOpenHighLessThan10','eanMSOpenHighGreaterThan10','eanMSOpenHighGreaterThan50','eanMSOpenHighGreaterThan100','eanMSOpenAverageLessThan10','eanMSOpenAverageGreaterThan10','eanMSOpenAverageGreaterThan50','eanMSOpenAverageGreaterThan100','eanMSOpenLowLessThan10','eanMSOpenLowGreaterThan10','eanMSOpenLowGreaterThan50','eanMSOpenLowGreaterThan100','eanMSPending_ClientCriticalLessThan10','eanMSPending_ClientCriticalGreaterThan10','eanMSPending_ClientCriticalGreaterThan50','eanMSPending_ClientCriticalGreaterThan100','eanMSPending_ClientHighLessThan10','eanMSPending_ClientHighGreaterThan10','eanMSPending_ClientHighGreaterThan50','eanMSPending_ClientHighGreaterThan100','eanMSPending_ClientAverageLessThan10','eanMSPending_ClientAverageGreaterThan10','eanMSPending_ClientAverageGreaterThan50','eanMSPending_ClientAverageGreaterThan100','eanMSPending_ClientLowLessThan10','eanMSPending_ClientLowGreaterThan10','eanMSPending_ClientLowGreaterThan50','eanMSPending_ClientLowGreaterThan100','eanMSPending_Emergency_ClientCriticalLessThan10','eanMSPending_Emergency_ClientCriticalGreaterThan10','eanMSPending_Emergency_ClientCriticalGreaterThan50','eanMSPending_Emergency_ClientCriticalGreaterThan100','eanMSPending_Emergency_ClientHighLessThan10','eanMSPending_Emergency_ClientHighGreaterThan10','eanMSPending_Emergency_ClientHighGreaterThan50','eanMSPending_Emergency_ClientHighGreaterThan100','eanMSPending_Emergency_ClientAverageLessThan10','eanMSPending_Emergency_ClientAverageGreaterThan10','eanMSPending_Emergency_ClientAverageGreaterThan50','eanMSPending_Emergency_ClientAverageGreaterThan100','eanMSPending_Emergency_ClientLowLessThan10','eanMSPending_Emergency_ClientLowGreaterThan10','eanMSPending_Emergency_ClientLowGreaterThan50','eanMSPending_Emergency_ClientLowGreaterThan100','eanMSPending_VendorCriticalLessThan10','eanMSPending_VendorCriticalGreaterThan10','eanMSPending_VendorCriticalGreaterThan50','eanMSPending_VendorCriticalGreaterThan100','eanMSPending_VendorHighLessThan10','eanMSPending_VendorHighGreaterThan10','eanMSPending_VendorHighGreaterThan50','eanMSPending_VendorHighGreaterThan100','eanMSPending_VendorAverageLessThan10','eanMSPending_VendorAverageGreaterThan10','eanMSPending_VendorAverageGreaterThan50','eanMSPending_VendorAverageGreaterThan100','eanMSPending_VendorLowLessThan10','eanMSPending_VendorLowGreaterThan10','eanMSPending_VendorLowGreaterThan50','eanMSPending_VendorLowGreaterThan100','eanMSWork_In_ProgressCriticalLessThan10','eanMSWork_In_ProgressCriticalGreaterThan10','eanMSWork_In_ProgressCriticalGreaterThan50','eanMSWork_In_ProgressCriticalGreaterThan100','eanMSWork_In_ProgressHighLessThan10','eanMSWork_In_ProgressHighGreaterThan10','eanMSWork_In_ProgressHighGreaterThan50','eanMSWork_In_ProgressHighGreaterThan100','eanMSWork_In_ProgressAverageLessThan10','eanMSWork_In_ProgressAverageGreaterThan10','eanMSWork_In_ProgressAverageGreaterThan50','eanMSWork_In_ProgressAverageGreaterThan100','eanMSWork_In_ProgressLowLessThan10','eanMSWork_In_ProgressLowGreaterThan10','eanMSWork_In_ProgressLowGreaterThan50','eanMSWork_In_ProgressLowGreaterThan100','eanNonMSOpenCriticalLessThan10','eanNonMSOpenCriticalGreaterThan10','eanNonMSOpenCriticalGreaterThan50','eanNonMSOpenCriticalGreaterThan100','eanNonMSOpenHighLessThan10','eanNonMSOpenHighGreaterThan10','eanNonMSOpenHighGreaterThan50','eanNonMSOpenHighGreaterThan100','eanNonMSOpenAverageLessThan10','eanNonMSOpenAverageGreaterThan10','eanNonMSOpenAverageGreaterThan50','eanNonMSOpenAverageGreaterThan100','eanNonMSOpenLowLessThan10','eanNonMSOpenLowGreaterThan10','eanNonMSOpenLowGreaterThan50','eanNonMSOpenLowGreaterThan100','eanNonMSPending_ClientCriticalLessThan10','eanNonMSPending_ClientCriticalGreaterThan10','eanNonMSPending_ClientCriticalGreaterThan50','eanNonMSPending_ClientCriticalGreaterThan100','eanNonMSPending_ClientHighLessThan10','eanNonMSPending_ClientHighGreaterThan10','eanNonMSPending_ClientHighGreaterThan50','eanNonMSPending_ClientHighGreaterThan100','eanNonMSPending_ClientAverageLessThan10','eanNonMSPending_ClientAverageGreaterThan10','eanNonMSPending_ClientAverageGreaterThan50','eanNonMSPending_ClientAverageGreaterThan100','eanNonMSPending_ClientLowLessThan10','eanNonMSPending_ClientLowGreaterThan10','eanNonMSPending_ClientLowGreaterThan50','eanNonMSPending_ClientLowGreaterThan100','eanNonMSPending_Emergency_ClientCriticalLessThan10','eanNonMSPending_Emergency_ClientCriticalGreaterThan10','eanNonMSPending_Emergency_ClientCriticalGreaterThan50','eanNonMSPending_Emergency_ClientCriticalGreaterThan100','eanNonMSPending_Emergency_ClientHighLessThan10','eanNonMSPending_Emergency_ClientHighGreaterThan10','eanNonMSPending_Emergency_ClientHighGreaterThan50','eanNonMSPending_Emergency_ClientHighGreaterThan100','eanNonMSPending_Emergency_ClientAverageLessThan10','eanNonMSPending_Emergency_ClientAverageGreaterThan10','eanNonMSPending_Emergency_ClientAverageGreaterThan50','eanNonMSPending_Emergency_ClientAverageGreaterThan100','eanNonMSPending_Emergency_ClientLowLessThan10','eanNonMSPending_Emergency_ClientLowGreaterThan10','eanNonMSPending_Emergency_ClientLowGreaterThan50','eanNonMSPending_Emergency_ClientLowGreaterThan100','eanNonMSPending_VendorCriticalLessThan10','eanNonMSPending_VendorCriticalGreaterThan10','eanNonMSPending_VendorCriticalGreaterThan50','eanNonMSPending_VendorCriticalGreaterThan100','eanNonMSPending_VendorHighLessThan10','eanNonMSPending_VendorHighGreaterThan10','eanNonMSPending_VendorHighGreaterThan50','eanNonMSPending_VendorHighGreaterThan100','eanNonMSPending_VendorAverageLessThan10','eanNonMSPending_VendorAverageGreaterThan10','eanNonMSPending_VendorAverageGreaterThan50','eanNonMSPending_VendorAverageGreaterThan100','eanNonMSPending_VendorLowLessThan10','eanNonMSPending_VendorLowGreaterThan10','eanNonMSPending_VendorLowGreaterThan50','eanNonMSPending_VendorLowGreaterThan100','eanNonMSWork_In_ProgressCriticalLessThan10','eanNonMSWork_In_ProgressCriticalGreaterThan10','eanNonMSWork_In_ProgressCriticalGreaterThan50','eanNonMSWork_In_ProgressCriticalGreaterThan100','eanNonMSWork_In_ProgressHighLessThan10','eanNonMSWork_In_ProgressHighGreaterThan10','eanNonMSWork_In_ProgressHighGreaterThan50','eanNonMSWork_In_ProgressHighGreaterThan100','eanNonMSWork_In_ProgressAverageLessThan10','eanNonMSWork_In_ProgressAverageGreaterThan10','eanNonMSWork_In_ProgressAverageGreaterThan50','eanNonMSWork_In_ProgressAverageGreaterThan100','eanNonMSWork_In_ProgressLowLessThan10','eanNonMSWork_In_ProgressLowGreaterThan10','eanNonMSWork_In_ProgressLowGreaterThan50','eanNonMSWork_In_ProgressLowGreaterThan100'
],'RowCount':[eanMSOpenCriticalLessThan10.shape[0],eanMSOpenCriticalGreaterThan10.shape[0],eanMSOpenCriticalGreaterThan50.shape[0],eanMSOpenCriticalGreaterThan100.shape[0],eanMSOpenHighLessThan10.shape[0],eanMSOpenHighGreaterThan10.shape[0],eanMSOpenHighGreaterThan50.shape[0],eanMSOpenHighGreaterThan100.shape[0],eanMSOpenAverageLessThan10.shape[0],eanMSOpenAverageGreaterThan10.shape[0],eanMSOpenAverageGreaterThan50.shape[0],eanMSOpenAverageGreaterThan100.shape[0],eanMSOpenLowLessThan10.shape[0],eanMSOpenLowGreaterThan10.shape[0],eanMSOpenLowGreaterThan50.shape[0],eanMSOpenLowGreaterThan100.shape[0],eanMSPending_ClientCriticalLessThan10.shape[0],eanMSPending_ClientCriticalGreaterThan10.shape[0],eanMSPending_ClientCriticalGreaterThan50.shape[0],eanMSPending_ClientCriticalGreaterThan100.shape[0],eanMSPending_ClientHighLessThan10.shape[0],eanMSPending_ClientHighGreaterThan10.shape[0],eanMSPending_ClientHighGreaterThan50.shape[0],eanMSPending_ClientHighGreaterThan100.shape[0],eanMSPending_ClientAverageLessThan10.shape[0],eanMSPending_ClientAverageGreaterThan10.shape[0],eanMSPending_ClientAverageGreaterThan50.shape[0],eanMSPending_ClientAverageGreaterThan100.shape[0],eanMSPending_ClientLowLessThan10.shape[0],eanMSPending_ClientLowGreaterThan10.shape[0],eanMSPending_ClientLowGreaterThan50.shape[0],eanMSPending_ClientLowGreaterThan100.shape[0],eanMSPending_Emergency_ClientCriticalLessThan10.shape[0],eanMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],eanMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],eanMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],eanMSPending_Emergency_ClientHighLessThan10.shape[0],eanMSPending_Emergency_ClientHighGreaterThan10.shape[0],eanMSPending_Emergency_ClientHighGreaterThan50.shape[0],eanMSPending_Emergency_ClientHighGreaterThan100.shape[0],eanMSPending_Emergency_ClientAverageLessThan10.shape[0],eanMSPending_Emergency_ClientAverageGreaterThan10.shape[0],eanMSPending_Emergency_ClientAverageGreaterThan50.shape[0],eanMSPending_Emergency_ClientAverageGreaterThan100.shape[0],eanMSPending_Emergency_ClientLowLessThan10.shape[0],eanMSPending_Emergency_ClientLowGreaterThan10.shape[0],eanMSPending_Emergency_ClientLowGreaterThan50.shape[0],eanMSPending_Emergency_ClientLowGreaterThan100.shape[0],eanMSPending_VendorCriticalLessThan10.shape[0],eanMSPending_VendorCriticalGreaterThan10.shape[0],eanMSPending_VendorCriticalGreaterThan50.shape[0],eanMSPending_VendorCriticalGreaterThan100.shape[0],eanMSPending_VendorHighLessThan10.shape[0],eanMSPending_VendorHighGreaterThan10.shape[0],eanMSPending_VendorHighGreaterThan50.shape[0],eanMSPending_VendorHighGreaterThan100.shape[0],eanMSPending_VendorAverageLessThan10.shape[0],eanMSPending_VendorAverageGreaterThan10.shape[0],eanMSPending_VendorAverageGreaterThan50.shape[0],eanMSPending_VendorAverageGreaterThan100.shape[0],eanMSPending_VendorLowLessThan10.shape[0],eanMSPending_VendorLowGreaterThan10.shape[0],eanMSPending_VendorLowGreaterThan50.shape[0],eanMSPending_VendorLowGreaterThan100.shape[0],eanMSWork_In_ProgressCriticalLessThan10.shape[0],eanMSWork_In_ProgressCriticalGreaterThan10.shape[0],eanMSWork_In_ProgressCriticalGreaterThan50.shape[0],eanMSWork_In_ProgressCriticalGreaterThan100.shape[0],eanMSWork_In_ProgressHighLessThan10.shape[0],eanMSWork_In_ProgressHighGreaterThan10.shape[0],eanMSWork_In_ProgressHighGreaterThan50.shape[0],eanMSWork_In_ProgressHighGreaterThan100.shape[0],eanMSWork_In_ProgressAverageLessThan10.shape[0],eanMSWork_In_ProgressAverageGreaterThan10.shape[0],eanMSWork_In_ProgressAverageGreaterThan50.shape[0],eanMSWork_In_ProgressAverageGreaterThan100.shape[0],eanMSWork_In_ProgressLowLessThan10.shape[0],eanMSWork_In_ProgressLowGreaterThan10.shape[0],eanMSWork_In_ProgressLowGreaterThan50.shape[0],eanMSWork_In_ProgressLowGreaterThan100.shape[0],eanNonMSOpenCriticalLessThan10.shape[0],eanNonMSOpenCriticalGreaterThan10.shape[0],eanNonMSOpenCriticalGreaterThan50.shape[0],eanNonMSOpenCriticalGreaterThan100.shape[0],eanNonMSOpenHighLessThan10.shape[0],eanNonMSOpenHighGreaterThan10.shape[0],eanNonMSOpenHighGreaterThan50.shape[0],eanNonMSOpenHighGreaterThan100.shape[0],eanNonMSOpenAverageLessThan10.shape[0],eanNonMSOpenAverageGreaterThan10.shape[0],eanNonMSOpenAverageGreaterThan50.shape[0],eanNonMSOpenAverageGreaterThan100.shape[0],eanNonMSOpenLowLessThan10.shape[0],eanNonMSOpenLowGreaterThan10.shape[0],eanNonMSOpenLowGreaterThan50.shape[0],eanNonMSOpenLowGreaterThan100.shape[0],eanNonMSPending_ClientCriticalLessThan10.shape[0],eanNonMSPending_ClientCriticalGreaterThan10.shape[0],eanNonMSPending_ClientCriticalGreaterThan50.shape[0],eanNonMSPending_ClientCriticalGreaterThan100.shape[0],eanNonMSPending_ClientHighLessThan10.shape[0],eanNonMSPending_ClientHighGreaterThan10.shape[0],eanNonMSPending_ClientHighGreaterThan50.shape[0],eanNonMSPending_ClientHighGreaterThan100.shape[0],eanNonMSPending_ClientAverageLessThan10.shape[0],eanNonMSPending_ClientAverageGreaterThan10.shape[0],eanNonMSPending_ClientAverageGreaterThan50.shape[0],eanNonMSPending_ClientAverageGreaterThan100.shape[0],eanNonMSPending_ClientLowLessThan10.shape[0],eanNonMSPending_ClientLowGreaterThan10.shape[0],eanNonMSPending_ClientLowGreaterThan50.shape[0],eanNonMSPending_ClientLowGreaterThan100.shape[0],eanNonMSPending_Emergency_ClientCriticalLessThan10.shape[0],eanNonMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],eanNonMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],eanNonMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],eanNonMSPending_Emergency_ClientHighLessThan10.shape[0],eanNonMSPending_Emergency_ClientHighGreaterThan10.shape[0],eanNonMSPending_Emergency_ClientHighGreaterThan50.shape[0],eanNonMSPending_Emergency_ClientHighGreaterThan100.shape[0],eanNonMSPending_Emergency_ClientAverageLessThan10.shape[0],eanNonMSPending_Emergency_ClientAverageGreaterThan10.shape[0],eanNonMSPending_Emergency_ClientAverageGreaterThan50.shape[0],eanNonMSPending_Emergency_ClientAverageGreaterThan100.shape[0],eanNonMSPending_Emergency_ClientLowLessThan10.shape[0],eanNonMSPending_Emergency_ClientLowGreaterThan10.shape[0],eanNonMSPending_Emergency_ClientLowGreaterThan50.shape[0],eanNonMSPending_Emergency_ClientLowGreaterThan100.shape[0],eanNonMSPending_VendorCriticalLessThan10.shape[0],eanNonMSPending_VendorCriticalGreaterThan10.shape[0],eanNonMSPending_VendorCriticalGreaterThan50.shape[0],eanNonMSPending_VendorCriticalGreaterThan100.shape[0],eanNonMSPending_VendorHighLessThan10.shape[0],eanNonMSPending_VendorHighGreaterThan10.shape[0],eanNonMSPending_VendorHighGreaterThan50.shape[0],eanNonMSPending_VendorHighGreaterThan100.shape[0],eanNonMSPending_VendorAverageLessThan10.shape[0],eanNonMSPending_VendorAverageGreaterThan10.shape[0],eanNonMSPending_VendorAverageGreaterThan50.shape[0],eanNonMSPending_VendorAverageGreaterThan100.shape[0],eanNonMSPending_VendorLowLessThan10.shape[0],eanNonMSPending_VendorLowGreaterThan10.shape[0],eanNonMSPending_VendorLowGreaterThan50.shape[0],eanNonMSPending_VendorLowGreaterThan100.shape[0],eanNonMSWork_In_ProgressCriticalLessThan10.shape[0],eanNonMSWork_In_ProgressCriticalGreaterThan10.shape[0],eanNonMSWork_In_ProgressCriticalGreaterThan50.shape[0],eanNonMSWork_In_ProgressCriticalGreaterThan100.shape[0],eanNonMSWork_In_ProgressHighLessThan10.shape[0],eanNonMSWork_In_ProgressHighGreaterThan10.shape[0],eanNonMSWork_In_ProgressHighGreaterThan50.shape[0],eanNonMSWork_In_ProgressHighGreaterThan100.shape[0],eanNonMSWork_In_ProgressAverageLessThan10.shape[0],eanNonMSWork_In_ProgressAverageGreaterThan10.shape[0],eanNonMSWork_In_ProgressAverageGreaterThan50.shape[0],eanNonMSWork_In_ProgressAverageGreaterThan100.shape[0],eanNonMSWork_In_ProgressLowLessThan10.shape[0],eanNonMSWork_In_ProgressLowGreaterThan10.shape[0],eanNonMSWork_In_ProgressLowGreaterThan50.shape[0],eanNonMSWork_In_ProgressLowGreaterThan100.shape[0]
]}
rcdf=pd.DataFrame(rowcount)
rcdf.to_excel(rowpath,index=False)
print("success")

#rowcount
#eanMSOpenCriticalLessThan10rc=eanMSOpenCriticalLessThan10.shape[0]
##(eanMSWork_In_ProgressHighLessThan10)
##(Msean.shape[0])
##(eanMSOpen)
##(eanMSOpenCritical)













#writer = pd.ExcelWriter(excelPath.replace(".xlsx","1.xlsx"), engine='xlsxwriter')

#ean1.to_excel(writer, sheet_name=ng,index=False)
#eanpport1.to_excel(writer, sheet_name=es,index=False)
#ediAnalysts1.to_excel(writer, sheet_name=ea,index=False)
#webMethods1.to_excel(writer, sheet_name=we,index=False)

#writer.save()
##("success")
