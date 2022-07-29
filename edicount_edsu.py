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
edsu=pd.read_excel(excelPath,sheet_name=0,index=False,engine='openpyxl')

#MS_NONMS
Msedsu=edsu[edsu['MS_NONMS'] == "MS"]
NonMsedsu=edsu[edsu['MS_NONMS'] == "Non_ms"]
#MS Status
edsuMSOpen=Msedsu[Msedsu['Status'] == "Open"]
edsuMSPending_Client=Msedsu[Msedsu['Status'] == "Pending Client"]
edsuMSPending_Emergency_Client=Msedsu[Msedsu['Status'] == "Pending Emergency Client"]
edsuMSPending_Vendor=Msedsu[Msedsu['Status'] == "Pending Vendor"]
edsuMSWork_In_Progress=Msedsu[Msedsu['Status'] == "Work In Progress"]
#(edsuMSWork_In_Progress)
#Non_MS Status
edsuNonMSOpen=NonMsedsu[NonMsedsu['Status'] == "Open"]
edsuNonMSPending_Client=NonMsedsu[NonMsedsu['Status'] == "Pending Client"]
edsuNonMSPending_Emergency_Client=NonMsedsu[NonMsedsu['Status'] == "Pending Emergency Client"]
edsuNonMSPending_Vendor=NonMsedsu[NonMsedsu['Status'] == "Pending Vendor"]
edsuNonMSWork_In_Progress=NonMsedsu[NonMsedsu['Status'] == "Work In Progress"]

#MS_OPEN_PRIORITY**
edsuMSOpenCritical=edsuMSOpen[edsuMSOpen['Priority'] == 1]
edsuMSOpenHigh=edsuMSOpen[edsuMSOpen['Priority'] == 2]
edsuMSOpenAverage=edsuMSOpen[edsuMSOpen['Priority'] == 3]
edsuMSOpenLow=edsuMSOpen[edsuMSOpen['Priority'] == 4]

##MS_Pending_Client_PRIORITY
edsuMSPending_ClientCritical=edsuMSPending_Client[edsuMSPending_Client['Priority'] == 1]
edsuMSPending_ClientHigh=edsuMSPending_Client[edsuMSPending_Client['Priority'] == 2]
edsuMSPending_ClientAverage=edsuMSPending_Client[edsuMSPending_Client['Priority'] == 3]
edsuMSPending_ClientLow=edsuMSPending_Client[edsuMSPending_Client['Priority'] == 4]

#MS_Pending_Emergency_Client
edsuMSPending_Emergency_ClientCritical=edsuMSPending_Emergency_Client[edsuMSPending_Emergency_Client['Priority'] == 1]
edsuMSPending_Emergency_ClientHigh=edsuMSPending_Emergency_Client[edsuMSPending_Emergency_Client['Priority'] == 2]
edsuMSPending_Emergency_ClientAverage=edsuMSPending_Emergency_Client[edsuMSPending_Emergency_Client['Priority'] == 3]
edsuMSPending_Emergency_ClientLow=edsuMSPending_Emergency_Client[edsuMSPending_Emergency_Client['Priority'] == 4]

#MS_Pending_Vendor
edsuMSPending_VendorCritical=edsuMSPending_Vendor[edsuMSPending_Vendor['Priority'] == 1]
edsuMSPending_VendorHigh=edsuMSPending_Vendor[edsuMSPending_Vendor['Priority'] == 2]
edsuMSPending_VendorAverage=edsuMSPending_Vendor[edsuMSPending_Vendor['Priority'] == 3]
edsuMSPending_VendorLow=edsuMSPending_Vendor[edsuMSPending_Vendor['Priority'] == 4]

#MS_Work_In_Progress
edsuMSWork_In_ProgressCritical=edsuMSWork_In_Progress[edsuMSWork_In_Progress['Priority'] == 1]
edsuMSWork_In_ProgressHigh=edsuMSWork_In_Progress[edsuMSWork_In_Progress['Priority'] == 2]
edsuMSWork_In_ProgressAverage=edsuMSWork_In_Progress[edsuMSWork_In_Progress['Priority'] == 3]
edsuMSWork_In_ProgressLow=edsuMSWork_In_Progress[edsuMSWork_In_Progress['Priority'] == 4]
#(edsuMSWork_In_ProgressHigh)


#NonMS_OPEN_PRIORITY
edsuNonMSOpenCritical=edsuNonMSOpen[edsuNonMSOpen['Priority'] == 1]
edsuNonMSOpenHigh=edsuNonMSOpen[edsuNonMSOpen['Priority'] == 2]
edsuNonMSOpenAverage=edsuNonMSOpen[edsuNonMSOpen['Priority'] == 3]
edsuNonMSOpenLow=edsuNonMSOpen[edsuNonMSOpen['Priority'] == 4]

##NonMS_Pending_Client_PRIORITY
edsuNonMSPending_ClientCritical=edsuNonMSPending_Client[edsuNonMSPending_Client['Priority'] == 1]
edsuNonMSPending_ClientHigh=edsuNonMSPending_Client[edsuNonMSPending_Client['Priority'] == 2]
edsuNonMSPending_ClientAverage=edsuNonMSPending_Client[edsuNonMSPending_Client['Priority'] == 3]
edsuNonMSPending_ClientLow=edsuNonMSPending_Client[edsuNonMSPending_Client['Priority'] == 4]

#NonMS_Pending_Emergency_Client
edsuNonMSPending_Emergency_ClientCritical=edsuNonMSPending_Emergency_Client[edsuNonMSPending_Emergency_Client['Priority'] == 1]
edsuNonMSPending_Emergency_ClientHigh=edsuNonMSPending_Emergency_Client[edsuNonMSPending_Emergency_Client['Priority'] == 2]
edsuNonMSPending_Emergency_ClientAverage=edsuNonMSPending_Emergency_Client[edsuNonMSPending_Emergency_Client['Priority'] == 3]
edsuNonMSPending_Emergency_ClientLow=edsuNonMSPending_Emergency_Client[edsuNonMSPending_Emergency_Client['Priority'] == 4]

#NonMS_Pending_Vendor
edsuNonMSPending_VendorCritical=edsuNonMSPending_Vendor[edsuNonMSPending_Vendor['Priority'] == 1]
edsuNonMSPending_VendorHigh=edsuNonMSPending_Vendor[edsuNonMSPending_Vendor['Priority'] == 2]
edsuNonMSPending_VendorAverage=edsuNonMSPending_Vendor[edsuNonMSPending_Vendor['Priority'] == 3]
edsuNonMSPending_VendorLow=edsuNonMSPending_Vendor[edsuNonMSPending_Vendor['Priority'] == 4]

#NonMS_Work_In_Progress
edsuNonMSWork_In_ProgressCritical=edsuNonMSWork_In_Progress[edsuNonMSWork_In_Progress['Priority'] == 1]
edsuNonMSWork_In_ProgressHigh=edsuNonMSWork_In_Progress[edsuNonMSWork_In_Progress['Priority'] == 2]
edsuNonMSWork_In_ProgressAverage=edsuNonMSWork_In_Progress[edsuNonMSWork_In_Progress['Priority'] == 3]
edsuNonMSWork_In_ProgressLow=edsuNonMSWork_In_Progress[edsuNonMSWork_In_Progress['Priority'] == 4]

#edsuMSOpenCriticalAging
edsuMSOpenCriticalLessThan10=edsuMSOpenCritical[edsuMSOpenCritical['Ageing days'] <=10]
edsuMSOpenCriticalGreaterThan10=edsuMSOpenCritical[edsuMSOpenCritical['Ageing days'] > 10]
edsuMSOpenCriticalGreaterThan50=edsuMSOpenCritical[edsuMSOpenCritical['Ageing days'] > 50]
edsuMSOpenCriticalGreaterThan100=edsuMSOpenCritical[edsuMSOpenCritical['Ageing days'] > 100]

#edsuMSOpenHighAging
edsuMSOpenHighLessThan10=edsuMSOpenHigh[edsuMSOpenHigh['Ageing days'] <=10]
edsuMSOpenHighGreaterThan10=edsuMSOpenHigh[edsuMSOpenHigh['Ageing days'] > 10]
edsuMSOpenHighGreaterThan50=edsuMSOpenHigh[edsuMSOpenHigh['Ageing days'] > 50]
edsuMSOpenHighGreaterThan100=edsuMSOpenHigh[edsuMSOpenHigh['Ageing days'] > 100]

#edsuMSOpenAverageAging
edsuMSOpenAverageLessThan10=edsuMSOpenAverage[edsuMSOpenAverage['Ageing days'] <=10]
edsuMSOpenAverageGreaterThan10=edsuMSOpenAverage[edsuMSOpenAverage['Ageing days'] > 10]
edsuMSOpenAverageGreaterThan50=edsuMSOpenAverage[edsuMSOpenAverage['Ageing days'] > 50]
edsuMSOpenAverageGreaterThan100=edsuMSOpenAverage[edsuMSOpenAverage['Ageing days'] > 100]

#edsuMSOpenLowAging
edsuMSOpenLowLessThan10=edsuMSOpenLow[edsuMSOpenLow['Ageing days'] <=10]
edsuMSOpenLowGreaterThan10=edsuMSOpenLow[edsuMSOpenLow['Ageing days'] > 10]
edsuMSOpenLowGreaterThan50=edsuMSOpenLow[edsuMSOpenLow['Ageing days'] > 50]
edsuMSOpenLowGreaterThan100=edsuMSOpenLow[edsuMSOpenLow['Ageing days'] > 100]


#edsuMSPending_ClientCriticalAging
edsuMSPending_ClientCriticalLessThan10=edsuMSPending_ClientCritical[edsuMSPending_ClientCritical['Ageing days'] <=10]
edsuMSPending_ClientCriticalGreaterThan10=edsuMSPending_ClientCritical[edsuMSPending_ClientCritical['Ageing days'] > 10]
edsuMSPending_ClientCriticalGreaterThan50=edsuMSPending_ClientCritical[edsuMSPending_ClientCritical['Ageing days'] > 50]
edsuMSPending_ClientCriticalGreaterThan100=edsuMSPending_ClientCritical[edsuMSPending_ClientCritical['Ageing days'] > 100]

#edsuMSPending_ClientHighAging
edsuMSPending_ClientHighLessThan10=edsuMSPending_ClientHigh[edsuMSPending_ClientHigh['Ageing days'] <=10]
edsuMSPending_ClientHighGreaterThan10=edsuMSPending_ClientHigh[edsuMSPending_ClientHigh['Ageing days'] > 10]
edsuMSPending_ClientHighGreaterThan50=edsuMSPending_ClientHigh[edsuMSPending_ClientHigh['Ageing days'] > 50]
edsuMSPending_ClientHighGreaterThan100=edsuMSPending_ClientHigh[edsuMSPending_ClientHigh['Ageing days'] > 100]

#edsuMSPending_ClientAverageAging
edsuMSPending_ClientAverageLessThan10=edsuMSPending_ClientAverage[edsuMSPending_ClientAverage['Ageing days'] <=10]
edsuMSPending_ClientAverageGreaterThan10=edsuMSPending_ClientAverage[edsuMSPending_ClientAverage['Ageing days'] > 10]
edsuMSPending_ClientAverageGreaterThan50=edsuMSPending_ClientAverage[edsuMSPending_ClientAverage['Ageing days'] > 50]
edsuMSPending_ClientAverageGreaterThan100=edsuMSPending_ClientAverage[edsuMSPending_ClientAverage['Ageing days'] > 100]

#edsuMSPending_ClientLowAging
edsuMSPending_ClientLowLessThan10=edsuMSPending_ClientLow[edsuMSPending_ClientLow['Ageing days'] <=10]
edsuMSPending_ClientLowGreaterThan10=edsuMSPending_ClientLow[edsuMSPending_ClientLow['Ageing days'] > 10]
edsuMSPending_ClientLowGreaterThan50=edsuMSPending_ClientLow[edsuMSPending_ClientLow['Ageing days'] > 50]
edsuMSPending_ClientLowGreaterThan100=edsuMSPending_ClientLow[edsuMSPending_ClientLow['Ageing days'] > 100]


#edsuMSPending_Emergency_ClientCriticalAging
edsuMSPending_Emergency_ClientCriticalLessThan10=edsuMSPending_Emergency_ClientCritical[edsuMSPending_Emergency_ClientCritical['Ageing days'] <=10]
edsuMSPending_Emergency_ClientCriticalGreaterThan10=edsuMSPending_Emergency_ClientCritical[edsuMSPending_Emergency_ClientCritical['Ageing days'] > 10]
edsuMSPending_Emergency_ClientCriticalGreaterThan50=edsuMSPending_Emergency_ClientCritical[edsuMSPending_Emergency_ClientCritical['Ageing days'] > 50]
edsuMSPending_Emergency_ClientCriticalGreaterThan100=edsuMSPending_Emergency_ClientCritical[edsuMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#edsuMSPending_Emergency_ClientHighAging
edsuMSPending_Emergency_ClientHighLessThan10=edsuMSPending_Emergency_ClientHigh[edsuMSPending_Emergency_ClientHigh['Ageing days'] <=10]
edsuMSPending_Emergency_ClientHighGreaterThan10=edsuMSPending_Emergency_ClientHigh[edsuMSPending_Emergency_ClientHigh['Ageing days'] > 10]
edsuMSPending_Emergency_ClientHighGreaterThan50=edsuMSPending_Emergency_ClientHigh[edsuMSPending_Emergency_ClientHigh['Ageing days'] > 50]
edsuMSPending_Emergency_ClientHighGreaterThan100=edsuMSPending_Emergency_ClientHigh[edsuMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#edsuMSPending_Emergency_ClientAverageAging
edsuMSPending_Emergency_ClientAverageLessThan10=edsuMSPending_Emergency_ClientAverage[edsuMSPending_Emergency_ClientAverage['Ageing days'] <=10]
edsuMSPending_Emergency_ClientAverageGreaterThan10=edsuMSPending_Emergency_ClientAverage[edsuMSPending_Emergency_ClientAverage['Ageing days'] > 10]
edsuMSPending_Emergency_ClientAverageGreaterThan50=edsuMSPending_Emergency_ClientAverage[edsuMSPending_Emergency_ClientAverage['Ageing days'] > 50]
edsuMSPending_Emergency_ClientAverageGreaterThan100=edsuMSPending_Emergency_ClientAverage[edsuMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#edsuMSPending_Emergency_ClientLowAging
edsuMSPending_Emergency_ClientLowLessThan10=edsuMSPending_Emergency_ClientLow[edsuMSPending_Emergency_ClientLow['Ageing days'] <=10]
edsuMSPending_Emergency_ClientLowGreaterThan10=edsuMSPending_Emergency_ClientLow[edsuMSPending_Emergency_ClientLow['Ageing days'] > 10]
edsuMSPending_Emergency_ClientLowGreaterThan50=edsuMSPending_Emergency_ClientLow[edsuMSPending_Emergency_ClientLow['Ageing days'] > 50]
edsuMSPending_Emergency_ClientLowGreaterThan100=edsuMSPending_Emergency_ClientLow[edsuMSPending_Emergency_ClientLow['Ageing days'] > 100]


#edsuMSPending_VendorCriticalAging
edsuMSPending_VendorCriticalLessThan10=edsuMSPending_VendorCritical[edsuMSPending_VendorCritical['Ageing days'] <=10]
edsuMSPending_VendorCriticalGreaterThan10=edsuMSPending_VendorCritical[edsuMSPending_VendorCritical['Ageing days'] > 10]
edsuMSPending_VendorCriticalGreaterThan50=edsuMSPending_VendorCritical[edsuMSPending_VendorCritical['Ageing days'] > 50]
edsuMSPending_VendorCriticalGreaterThan100=edsuMSPending_VendorCritical[edsuMSPending_VendorCritical['Ageing days'] > 100]

#edsuMSPending_VendorHighAging
edsuMSPending_VendorHighLessThan10=edsuMSPending_VendorHigh[edsuMSPending_VendorHigh['Ageing days'] <=10]
edsuMSPending_VendorHighGreaterThan10=edsuMSPending_VendorHigh[edsuMSPending_VendorHigh['Ageing days'] > 10]
edsuMSPending_VendorHighGreaterThan50=edsuMSPending_VendorHigh[edsuMSPending_VendorHigh['Ageing days'] > 50]
edsuMSPending_VendorHighGreaterThan100=edsuMSPending_VendorHigh[edsuMSPending_VendorHigh['Ageing days'] > 100]

#edsuMSPending_VendorAverageAging
edsuMSPending_VendorAverageLessThan10=edsuMSPending_VendorAverage[edsuMSPending_VendorAverage['Ageing days'] <=10]
edsuMSPending_VendorAverageGreaterThan10=edsuMSPending_VendorAverage[edsuMSPending_VendorAverage['Ageing days'] > 10]
edsuMSPending_VendorAverageGreaterThan50=edsuMSPending_VendorAverage[edsuMSPending_VendorAverage['Ageing days'] > 50]
edsuMSPending_VendorAverageGreaterThan100=edsuMSPending_VendorAverage[edsuMSPending_VendorAverage['Ageing days'] > 100]

#edsuMSPending_VendorLowAging
edsuMSPending_VendorLowLessThan10=edsuMSPending_VendorLow[edsuMSPending_VendorLow['Ageing days'] <=10]
edsuMSPending_VendorLowGreaterThan10=edsuMSPending_VendorLow[edsuMSPending_VendorLow['Ageing days'] > 10]
edsuMSPending_VendorLowGreaterThan50=edsuMSPending_VendorLow[edsuMSPending_VendorLow['Ageing days'] > 50]
edsuMSPending_VendorLowGreaterThan100=edsuMSPending_VendorLow[edsuMSPending_VendorLow['Ageing days'] > 100]


#edsuMSWork_In_ProgressCriticalAging
edsuMSWork_In_ProgressCriticalLessThan10=edsuMSWork_In_ProgressCritical[edsuMSWork_In_ProgressCritical['Ageing days'] <=10]
edsuMSWork_In_ProgressCriticalGreaterThan10=edsuMSWork_In_ProgressCritical[edsuMSWork_In_ProgressCritical['Ageing days'] > 10]
edsuMSWork_In_ProgressCriticalGreaterThan50=edsuMSWork_In_ProgressCritical[edsuMSWork_In_ProgressCritical['Ageing days'] > 50]
edsuMSWork_In_ProgressCriticalGreaterThan100=edsuMSWork_In_ProgressCritical[edsuMSWork_In_ProgressCritical['Ageing days'] > 100]

#edsuMSWork_In_ProgressHighAging
edsuMSWork_In_ProgressHighLessThan10=edsuMSWork_In_ProgressHigh[edsuMSWork_In_ProgressHigh['Ageing days'] <=10]
edsuMSWork_In_ProgressHighGreaterThan10=edsuMSWork_In_ProgressHigh[edsuMSWork_In_ProgressHigh['Ageing days'] > 10]
edsuMSWork_In_ProgressHighGreaterThan50=edsuMSWork_In_ProgressHigh[edsuMSWork_In_ProgressHigh['Ageing days'] > 50]
edsuMSWork_In_ProgressHighGreaterThan100=edsuMSWork_In_ProgressHigh[edsuMSWork_In_ProgressHigh['Ageing days'] > 100]
#(edsuMSWork_In_ProgressHighLessThan10)
#(edsuMSWork_In_ProgressHighLessThan10.shape[0])
#edsuMSWork_In_ProgressAverageAging
edsuMSWork_In_ProgressAverageLessThan10=edsuMSWork_In_ProgressAverage[edsuMSWork_In_ProgressAverage['Ageing days'] <=10]
edsuMSWork_In_ProgressAverageGreaterThan10=edsuMSWork_In_ProgressAverage[edsuMSWork_In_ProgressAverage['Ageing days'] > 10]
edsuMSWork_In_ProgressAverageGreaterThan50=edsuMSWork_In_ProgressAverage[edsuMSWork_In_ProgressAverage['Ageing days'] > 50]
edsuMSWork_In_ProgressAverageGreaterThan100=edsuMSWork_In_ProgressAverage[edsuMSWork_In_ProgressAverage['Ageing days'] > 100]

#edsuMSWork_In_ProgressLowAging
edsuMSWork_In_ProgressLowLessThan10=edsuMSWork_In_ProgressLow[edsuMSWork_In_ProgressLow['Ageing days'] <=10]
edsuMSWork_In_ProgressLowGreaterThan10=edsuMSWork_In_ProgressLow[edsuMSWork_In_ProgressLow['Ageing days'] > 10]
edsuMSWork_In_ProgressLowGreaterThan50=edsuMSWork_In_ProgressLow[edsuMSWork_In_ProgressLow['Ageing days'] > 50]
edsuMSWork_In_ProgressLowGreaterThan100=edsuMSWork_In_ProgressLow[edsuMSWork_In_ProgressLow['Ageing days'] > 100]



#edsuNonMSOpenCriticalAging
edsuNonMSOpenCriticalLessThan10=edsuNonMSOpenCritical[edsuNonMSOpenCritical['Ageing days'] <=10]
edsuNonMSOpenCriticalGreaterThan10=edsuNonMSOpenCritical[edsuNonMSOpenCritical['Ageing days'] > 10]
edsuNonMSOpenCriticalGreaterThan50=edsuNonMSOpenCritical[edsuNonMSOpenCritical['Ageing days'] > 50]
edsuNonMSOpenCriticalGreaterThan100=edsuNonMSOpenCritical[edsuNonMSOpenCritical['Ageing days'] > 100]

#edsuNonMSOpenHighAging
edsuNonMSOpenHighLessThan10=edsuNonMSOpenHigh[edsuNonMSOpenHigh['Ageing days'] <=10]
edsuNonMSOpenHighGreaterThan10=edsuNonMSOpenHigh[edsuNonMSOpenHigh['Ageing days'] > 10]
edsuNonMSOpenHighGreaterThan50=edsuNonMSOpenHigh[edsuNonMSOpenHigh['Ageing days'] > 50]
edsuNonMSOpenHighGreaterThan100=edsuNonMSOpenHigh[edsuNonMSOpenHigh['Ageing days'] > 100]

#edsuNonMSOpenAverageAging
edsuNonMSOpenAverageLessThan10=edsuNonMSOpenAverage[edsuNonMSOpenAverage['Ageing days'] <=10]
edsuNonMSOpenAverageGreaterThan10=edsuNonMSOpenAverage[edsuNonMSOpenAverage['Ageing days'] > 10]
edsuNonMSOpenAverageGreaterThan50=edsuNonMSOpenAverage[edsuNonMSOpenAverage['Ageing days'] > 50]
edsuNonMSOpenAverageGreaterThan100=edsuNonMSOpenAverage[edsuNonMSOpenAverage['Ageing days'] > 100]

#edsuNonMSOpenLowAging
edsuNonMSOpenLowLessThan10=edsuNonMSOpenLow[edsuNonMSOpenLow['Ageing days'] <=10]
edsuNonMSOpenLowGreaterThan10=edsuNonMSOpenLow[edsuNonMSOpenLow['Ageing days'] > 10]
edsuNonMSOpenLowGreaterThan50=edsuNonMSOpenLow[edsuNonMSOpenLow['Ageing days'] > 50]
edsuNonMSOpenLowGreaterThan100=edsuNonMSOpenLow[edsuNonMSOpenLow['Ageing days'] > 100]


#edsuNonMSPending_ClientCriticalAging
edsuNonMSPending_ClientCriticalLessThan10=edsuNonMSPending_ClientCritical[edsuNonMSPending_ClientCritical['Ageing days'] <=10]
edsuNonMSPending_ClientCriticalGreaterThan10=edsuNonMSPending_ClientCritical[edsuNonMSPending_ClientCritical['Ageing days'] > 10]
edsuNonMSPending_ClientCriticalGreaterThan50=edsuNonMSPending_ClientCritical[edsuNonMSPending_ClientCritical['Ageing days'] > 50]
edsuNonMSPending_ClientCriticalGreaterThan100=edsuNonMSPending_ClientCritical[edsuNonMSPending_ClientCritical['Ageing days'] > 100]

#edsuNonMSPending_ClientHighAging
edsuNonMSPending_ClientHighLessThan10=edsuNonMSPending_ClientHigh[edsuNonMSPending_ClientHigh['Ageing days'] <=10]
edsuNonMSPending_ClientHighGreaterThan10=edsuNonMSPending_ClientHigh[edsuNonMSPending_ClientHigh['Ageing days'] > 10]
edsuNonMSPending_ClientHighGreaterThan50=edsuNonMSPending_ClientHigh[edsuNonMSPending_ClientHigh['Ageing days'] > 50]
edsuNonMSPending_ClientHighGreaterThan100=edsuNonMSPending_ClientHigh[edsuNonMSPending_ClientHigh['Ageing days'] > 100]

#edsuNonMSPending_ClientAverageAging
edsuNonMSPending_ClientAverageLessThan10=edsuNonMSPending_ClientAverage[edsuNonMSPending_ClientAverage['Ageing days'] <=10]
edsuNonMSPending_ClientAverageGreaterThan10=edsuNonMSPending_ClientAverage[edsuNonMSPending_ClientAverage['Ageing days'] > 10]
edsuNonMSPending_ClientAverageGreaterThan50=edsuNonMSPending_ClientAverage[edsuNonMSPending_ClientAverage['Ageing days'] > 50]
edsuNonMSPending_ClientAverageGreaterThan100=edsuNonMSPending_ClientAverage[edsuNonMSPending_ClientAverage['Ageing days'] > 100]

#edsuNonMSPending_ClientLowAging
edsuNonMSPending_ClientLowLessThan10=edsuNonMSPending_ClientLow[edsuNonMSPending_ClientLow['Ageing days'] <=10]
edsuNonMSPending_ClientLowGreaterThan10=edsuNonMSPending_ClientLow[edsuNonMSPending_ClientLow['Ageing days'] > 10]
edsuNonMSPending_ClientLowGreaterThan50=edsuNonMSPending_ClientLow[edsuNonMSPending_ClientLow['Ageing days'] > 50]
edsuNonMSPending_ClientLowGreaterThan100=edsuNonMSPending_ClientLow[edsuNonMSPending_ClientLow['Ageing days'] > 100]


#edsuNonMSPending_Emergency_ClientCriticalAging
edsuNonMSPending_Emergency_ClientCriticalLessThan10=edsuNonMSPending_Emergency_ClientCritical[edsuNonMSPending_Emergency_ClientCritical['Ageing days'] <=10]
edsuNonMSPending_Emergency_ClientCriticalGreaterThan10=edsuNonMSPending_Emergency_ClientCritical[edsuNonMSPending_Emergency_ClientCritical['Ageing days'] > 10]
edsuNonMSPending_Emergency_ClientCriticalGreaterThan50=edsuNonMSPending_Emergency_ClientCritical[edsuNonMSPending_Emergency_ClientCritical['Ageing days'] > 50]
edsuNonMSPending_Emergency_ClientCriticalGreaterThan100=edsuNonMSPending_Emergency_ClientCritical[edsuNonMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#edsuNonMSPending_Emergency_ClientHighAging
edsuNonMSPending_Emergency_ClientHighLessThan10=edsuNonMSPending_Emergency_ClientHigh[edsuNonMSPending_Emergency_ClientHigh['Ageing days'] <=10]
edsuNonMSPending_Emergency_ClientHighGreaterThan10=edsuNonMSPending_Emergency_ClientHigh[edsuNonMSPending_Emergency_ClientHigh['Ageing days'] > 10]
edsuNonMSPending_Emergency_ClientHighGreaterThan50=edsuNonMSPending_Emergency_ClientHigh[edsuNonMSPending_Emergency_ClientHigh['Ageing days'] > 50]
edsuNonMSPending_Emergency_ClientHighGreaterThan100=edsuNonMSPending_Emergency_ClientHigh[edsuNonMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#edsuNonMSPending_Emergency_ClientAverageAging
edsuNonMSPending_Emergency_ClientAverageLessThan10=edsuNonMSPending_Emergency_ClientAverage[edsuNonMSPending_Emergency_ClientAverage['Ageing days'] <=10]
edsuNonMSPending_Emergency_ClientAverageGreaterThan10=edsuNonMSPending_Emergency_ClientAverage[edsuNonMSPending_Emergency_ClientAverage['Ageing days'] > 10]
edsuNonMSPending_Emergency_ClientAverageGreaterThan50=edsuNonMSPending_Emergency_ClientAverage[edsuNonMSPending_Emergency_ClientAverage['Ageing days'] > 50]
edsuNonMSPending_Emergency_ClientAverageGreaterThan100=edsuNonMSPending_Emergency_ClientAverage[edsuNonMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#edsuNonMSPending_Emergency_ClientLowAging
edsuNonMSPending_Emergency_ClientLowLessThan10=edsuNonMSPending_Emergency_ClientLow[edsuNonMSPending_Emergency_ClientLow['Ageing days'] <=10]
edsuNonMSPending_Emergency_ClientLowGreaterThan10=edsuNonMSPending_Emergency_ClientLow[edsuNonMSPending_Emergency_ClientLow['Ageing days'] > 10]
edsuNonMSPending_Emergency_ClientLowGreaterThan50=edsuNonMSPending_Emergency_ClientLow[edsuNonMSPending_Emergency_ClientLow['Ageing days'] > 50]
edsuNonMSPending_Emergency_ClientLowGreaterThan100=edsuNonMSPending_Emergency_ClientLow[edsuNonMSPending_Emergency_ClientLow['Ageing days'] > 100]


#edsuNonMSPending_VendorCriticalAging
edsuNonMSPending_VendorCriticalLessThan10=edsuNonMSPending_VendorCritical[edsuNonMSPending_VendorCritical['Ageing days'] <=10]
edsuNonMSPending_VendorCriticalGreaterThan10=edsuNonMSPending_VendorCritical[edsuNonMSPending_VendorCritical['Ageing days'] > 10]
edsuNonMSPending_VendorCriticalGreaterThan50=edsuNonMSPending_VendorCritical[edsuNonMSPending_VendorCritical['Ageing days'] > 50]
edsuNonMSPending_VendorCriticalGreaterThan100=edsuNonMSPending_VendorCritical[edsuNonMSPending_VendorCritical['Ageing days'] > 100]

#edsuNonMSPending_VendorHighAging
edsuNonMSPending_VendorHighLessThan10=edsuNonMSPending_VendorHigh[edsuNonMSPending_VendorHigh['Ageing days'] <=10]
edsuNonMSPending_VendorHighGreaterThan10=edsuNonMSPending_VendorHigh[edsuNonMSPending_VendorHigh['Ageing days'] > 10]
edsuNonMSPending_VendorHighGreaterThan50=edsuNonMSPending_VendorHigh[edsuNonMSPending_VendorHigh['Ageing days'] > 50]
edsuNonMSPending_VendorHighGreaterThan100=edsuNonMSPending_VendorHigh[edsuNonMSPending_VendorHigh['Ageing days'] > 100]

#edsuNonMSPending_VendorAverageAging
edsuNonMSPending_VendorAverageLessThan10=edsuNonMSPending_VendorAverage[edsuNonMSPending_VendorAverage['Ageing days'] <=10]
edsuNonMSPending_VendorAverageGreaterThan10=edsuNonMSPending_VendorAverage[edsuNonMSPending_VendorAverage['Ageing days'] > 10]
edsuNonMSPending_VendorAverageGreaterThan50=edsuNonMSPending_VendorAverage[edsuNonMSPending_VendorAverage['Ageing days'] > 50]
edsuNonMSPending_VendorAverageGreaterThan100=edsuNonMSPending_VendorAverage[edsuNonMSPending_VendorAverage['Ageing days'] > 100]

#edsuNonMSPending_VendorLowAging
edsuNonMSPending_VendorLowLessThan10=edsuNonMSPending_VendorLow[edsuNonMSPending_VendorLow['Ageing days'] <=10]
edsuNonMSPending_VendorLowGreaterThan10=edsuNonMSPending_VendorLow[edsuNonMSPending_VendorLow['Ageing days'] > 10]
edsuNonMSPending_VendorLowGreaterThan50=edsuNonMSPending_VendorLow[edsuNonMSPending_VendorLow['Ageing days'] > 50]
edsuNonMSPending_VendorLowGreaterThan100=edsuNonMSPending_VendorLow[edsuNonMSPending_VendorLow['Ageing days'] > 100]


#edsuNonMSWork_In_ProgressCriticalAging
edsuNonMSWork_In_ProgressCriticalLessThan10=edsuNonMSWork_In_ProgressCritical[edsuNonMSWork_In_ProgressCritical['Ageing days'] <=10]
edsuNonMSWork_In_ProgressCriticalGreaterThan10=edsuNonMSWork_In_ProgressCritical[edsuNonMSWork_In_ProgressCritical['Ageing days'] > 10]
edsuNonMSWork_In_ProgressCriticalGreaterThan50=edsuNonMSWork_In_ProgressCritical[edsuNonMSWork_In_ProgressCritical['Ageing days'] > 50]
edsuNonMSWork_In_ProgressCriticalGreaterThan100=edsuNonMSWork_In_ProgressCritical[edsuNonMSWork_In_ProgressCritical['Ageing days'] > 100]

#edsuNonMSWork_In_ProgressHighAging
edsuNonMSWork_In_ProgressHighLessThan10=edsuNonMSWork_In_ProgressHigh[edsuNonMSWork_In_ProgressHigh['Ageing days'] <=10]
edsuNonMSWork_In_ProgressHighGreaterThan10=edsuNonMSWork_In_ProgressHigh[edsuNonMSWork_In_ProgressHigh['Ageing days'] > 10]
edsuNonMSWork_In_ProgressHighGreaterThan50=edsuNonMSWork_In_ProgressHigh[edsuNonMSWork_In_ProgressHigh['Ageing days'] > 50]
edsuNonMSWork_In_ProgressHighGreaterThan100=edsuNonMSWork_In_ProgressHigh[edsuNonMSWork_In_ProgressHigh['Ageing days'] > 100]

#edsuNonMSWork_In_ProgressAverageAging
edsuNonMSWork_In_ProgressAverageLessThan10=edsuNonMSWork_In_ProgressAverage[edsuNonMSWork_In_ProgressAverage['Ageing days'] <=10]
edsuNonMSWork_In_ProgressAverageGreaterThan10=edsuNonMSWork_In_ProgressAverage[edsuNonMSWork_In_ProgressAverage['Ageing days'] > 10]
edsuNonMSWork_In_ProgressAverageGreaterThan50=edsuNonMSWork_In_ProgressAverage[edsuNonMSWork_In_ProgressAverage['Ageing days'] > 50]
edsuNonMSWork_In_ProgressAverageGreaterThan100=edsuNonMSWork_In_ProgressAverage[edsuNonMSWork_In_ProgressAverage['Ageing days'] > 100]

#edsuNonMSWork_In_ProgressLowAging
edsuNonMSWork_In_ProgressLowLessThan10=edsuNonMSWork_In_ProgressLow[edsuNonMSWork_In_ProgressLow['Ageing days'] <=10]
edsuNonMSWork_In_ProgressLowGreaterThan10=edsuNonMSWork_In_ProgressLow[edsuNonMSWork_In_ProgressLow['Ageing days'] > 10]
edsuNonMSWork_In_ProgressLowGreaterThan50=edsuNonMSWork_In_ProgressLow[edsuNonMSWork_In_ProgressLow['Ageing days'] > 50]
edsuNonMSWork_In_ProgressLowGreaterThan100=edsuNonMSWork_In_ProgressLow[edsuNonMSWork_In_ProgressLow['Ageing days'] > 100]


rowcount={'Data':['edsuMSOpenCriticalLessThan10','edsuMSOpenCriticalGreaterThan10','edsuMSOpenCriticalGreaterThan50','edsuMSOpenCriticalGreaterThan100','edsuMSOpenHighLessThan10','edsuMSOpenHighGreaterThan10','edsuMSOpenHighGreaterThan50','edsuMSOpenHighGreaterThan100','edsuMSOpenAverageLessThan10','edsuMSOpenAverageGreaterThan10','edsuMSOpenAverageGreaterThan50','edsuMSOpenAverageGreaterThan100','edsuMSOpenLowLessThan10','edsuMSOpenLowGreaterThan10','edsuMSOpenLowGreaterThan50','edsuMSOpenLowGreaterThan100','edsuMSPending_ClientCriticalLessThan10','edsuMSPending_ClientCriticalGreaterThan10','edsuMSPending_ClientCriticalGreaterThan50','edsuMSPending_ClientCriticalGreaterThan100','edsuMSPending_ClientHighLessThan10','edsuMSPending_ClientHighGreaterThan10','edsuMSPending_ClientHighGreaterThan50','edsuMSPending_ClientHighGreaterThan100','edsuMSPending_ClientAverageLessThan10','edsuMSPending_ClientAverageGreaterThan10','edsuMSPending_ClientAverageGreaterThan50','edsuMSPending_ClientAverageGreaterThan100','edsuMSPending_ClientLowLessThan10','edsuMSPending_ClientLowGreaterThan10','edsuMSPending_ClientLowGreaterThan50','edsuMSPending_ClientLowGreaterThan100','edsuMSPending_Emergency_ClientCriticalLessThan10','edsuMSPending_Emergency_ClientCriticalGreaterThan10','edsuMSPending_Emergency_ClientCriticalGreaterThan50','edsuMSPending_Emergency_ClientCriticalGreaterThan100','edsuMSPending_Emergency_ClientHighLessThan10','edsuMSPending_Emergency_ClientHighGreaterThan10','edsuMSPending_Emergency_ClientHighGreaterThan50','edsuMSPending_Emergency_ClientHighGreaterThan100','edsuMSPending_Emergency_ClientAverageLessThan10','edsuMSPending_Emergency_ClientAverageGreaterThan10','edsuMSPending_Emergency_ClientAverageGreaterThan50','edsuMSPending_Emergency_ClientAverageGreaterThan100','edsuMSPending_Emergency_ClientLowLessThan10','edsuMSPending_Emergency_ClientLowGreaterThan10','edsuMSPending_Emergency_ClientLowGreaterThan50','edsuMSPending_Emergency_ClientLowGreaterThan100','edsuMSPending_VendorCriticalLessThan10','edsuMSPending_VendorCriticalGreaterThan10','edsuMSPending_VendorCriticalGreaterThan50','edsuMSPending_VendorCriticalGreaterThan100','edsuMSPending_VendorHighLessThan10','edsuMSPending_VendorHighGreaterThan10','edsuMSPending_VendorHighGreaterThan50','edsuMSPending_VendorHighGreaterThan100','edsuMSPending_VendorAverageLessThan10','edsuMSPending_VendorAverageGreaterThan10','edsuMSPending_VendorAverageGreaterThan50','edsuMSPending_VendorAverageGreaterThan100','edsuMSPending_VendorLowLessThan10','edsuMSPending_VendorLowGreaterThan10','edsuMSPending_VendorLowGreaterThan50','edsuMSPending_VendorLowGreaterThan100','edsuMSWork_In_ProgressCriticalLessThan10','edsuMSWork_In_ProgressCriticalGreaterThan10','edsuMSWork_In_ProgressCriticalGreaterThan50','edsuMSWork_In_ProgressCriticalGreaterThan100','edsuMSWork_In_ProgressHighLessThan10','edsuMSWork_In_ProgressHighGreaterThan10','edsuMSWork_In_ProgressHighGreaterThan50','edsuMSWork_In_ProgressHighGreaterThan100','edsuMSWork_In_ProgressAverageLessThan10','edsuMSWork_In_ProgressAverageGreaterThan10','edsuMSWork_In_ProgressAverageGreaterThan50','edsuMSWork_In_ProgressAverageGreaterThan100','edsuMSWork_In_ProgressLowLessThan10','edsuMSWork_In_ProgressLowGreaterThan10','edsuMSWork_In_ProgressLowGreaterThan50','edsuMSWork_In_ProgressLowGreaterThan100','edsuNonMSOpenCriticalLessThan10','edsuNonMSOpenCriticalGreaterThan10','edsuNonMSOpenCriticalGreaterThan50','edsuNonMSOpenCriticalGreaterThan100','edsuNonMSOpenHighLessThan10','edsuNonMSOpenHighGreaterThan10','edsuNonMSOpenHighGreaterThan50','edsuNonMSOpenHighGreaterThan100','edsuNonMSOpenAverageLessThan10','edsuNonMSOpenAverageGreaterThan10','edsuNonMSOpenAverageGreaterThan50','edsuNonMSOpenAverageGreaterThan100','edsuNonMSOpenLowLessThan10','edsuNonMSOpenLowGreaterThan10','edsuNonMSOpenLowGreaterThan50','edsuNonMSOpenLowGreaterThan100','edsuNonMSPending_ClientCriticalLessThan10','edsuNonMSPending_ClientCriticalGreaterThan10','edsuNonMSPending_ClientCriticalGreaterThan50','edsuNonMSPending_ClientCriticalGreaterThan100','edsuNonMSPending_ClientHighLessThan10','edsuNonMSPending_ClientHighGreaterThan10','edsuNonMSPending_ClientHighGreaterThan50','edsuNonMSPending_ClientHighGreaterThan100','edsuNonMSPending_ClientAverageLessThan10','edsuNonMSPending_ClientAverageGreaterThan10','edsuNonMSPending_ClientAverageGreaterThan50','edsuNonMSPending_ClientAverageGreaterThan100','edsuNonMSPending_ClientLowLessThan10','edsuNonMSPending_ClientLowGreaterThan10','edsuNonMSPending_ClientLowGreaterThan50','edsuNonMSPending_ClientLowGreaterThan100','edsuNonMSPending_Emergency_ClientCriticalLessThan10','edsuNonMSPending_Emergency_ClientCriticalGreaterThan10','edsuNonMSPending_Emergency_ClientCriticalGreaterThan50','edsuNonMSPending_Emergency_ClientCriticalGreaterThan100','edsuNonMSPending_Emergency_ClientHighLessThan10','edsuNonMSPending_Emergency_ClientHighGreaterThan10','edsuNonMSPending_Emergency_ClientHighGreaterThan50','edsuNonMSPending_Emergency_ClientHighGreaterThan100','edsuNonMSPending_Emergency_ClientAverageLessThan10','edsuNonMSPending_Emergency_ClientAverageGreaterThan10','edsuNonMSPending_Emergency_ClientAverageGreaterThan50','edsuNonMSPending_Emergency_ClientAverageGreaterThan100','edsuNonMSPending_Emergency_ClientLowLessThan10','edsuNonMSPending_Emergency_ClientLowGreaterThan10','edsuNonMSPending_Emergency_ClientLowGreaterThan50','edsuNonMSPending_Emergency_ClientLowGreaterThan100','edsuNonMSPending_VendorCriticalLessThan10','edsuNonMSPending_VendorCriticalGreaterThan10','edsuNonMSPending_VendorCriticalGreaterThan50','edsuNonMSPending_VendorCriticalGreaterThan100','edsuNonMSPending_VendorHighLessThan10','edsuNonMSPending_VendorHighGreaterThan10','edsuNonMSPending_VendorHighGreaterThan50','edsuNonMSPending_VendorHighGreaterThan100','edsuNonMSPending_VendorAverageLessThan10','edsuNonMSPending_VendorAverageGreaterThan10','edsuNonMSPending_VendorAverageGreaterThan50','edsuNonMSPending_VendorAverageGreaterThan100','edsuNonMSPending_VendorLowLessThan10','edsuNonMSPending_VendorLowGreaterThan10','edsuNonMSPending_VendorLowGreaterThan50','edsuNonMSPending_VendorLowGreaterThan100','edsuNonMSWork_In_ProgressCriticalLessThan10','edsuNonMSWork_In_ProgressCriticalGreaterThan10','edsuNonMSWork_In_ProgressCriticalGreaterThan50','edsuNonMSWork_In_ProgressCriticalGreaterThan100','edsuNonMSWork_In_ProgressHighLessThan10','edsuNonMSWork_In_ProgressHighGreaterThan10','edsuNonMSWork_In_ProgressHighGreaterThan50','edsuNonMSWork_In_ProgressHighGreaterThan100','edsuNonMSWork_In_ProgressAverageLessThan10','edsuNonMSWork_In_ProgressAverageGreaterThan10','edsuNonMSWork_In_ProgressAverageGreaterThan50','edsuNonMSWork_In_ProgressAverageGreaterThan100','edsuNonMSWork_In_ProgressLowLessThan10','edsuNonMSWork_In_ProgressLowGreaterThan10','edsuNonMSWork_In_ProgressLowGreaterThan50','edsuNonMSWork_In_ProgressLowGreaterThan100'
],'RowCount':[edsuMSOpenCriticalLessThan10.shape[0],edsuMSOpenCriticalGreaterThan10.shape[0],edsuMSOpenCriticalGreaterThan50.shape[0],edsuMSOpenCriticalGreaterThan100.shape[0],edsuMSOpenHighLessThan10.shape[0],edsuMSOpenHighGreaterThan10.shape[0],edsuMSOpenHighGreaterThan50.shape[0],edsuMSOpenHighGreaterThan100.shape[0],edsuMSOpenAverageLessThan10.shape[0],edsuMSOpenAverageGreaterThan10.shape[0],edsuMSOpenAverageGreaterThan50.shape[0],edsuMSOpenAverageGreaterThan100.shape[0],edsuMSOpenLowLessThan10.shape[0],edsuMSOpenLowGreaterThan10.shape[0],edsuMSOpenLowGreaterThan50.shape[0],edsuMSOpenLowGreaterThan100.shape[0],edsuMSPending_ClientCriticalLessThan10.shape[0],edsuMSPending_ClientCriticalGreaterThan10.shape[0],edsuMSPending_ClientCriticalGreaterThan50.shape[0],edsuMSPending_ClientCriticalGreaterThan100.shape[0],edsuMSPending_ClientHighLessThan10.shape[0],edsuMSPending_ClientHighGreaterThan10.shape[0],edsuMSPending_ClientHighGreaterThan50.shape[0],edsuMSPending_ClientHighGreaterThan100.shape[0],edsuMSPending_ClientAverageLessThan10.shape[0],edsuMSPending_ClientAverageGreaterThan10.shape[0],edsuMSPending_ClientAverageGreaterThan50.shape[0],edsuMSPending_ClientAverageGreaterThan100.shape[0],edsuMSPending_ClientLowLessThan10.shape[0],edsuMSPending_ClientLowGreaterThan10.shape[0],edsuMSPending_ClientLowGreaterThan50.shape[0],edsuMSPending_ClientLowGreaterThan100.shape[0],edsuMSPending_Emergency_ClientCriticalLessThan10.shape[0],edsuMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],edsuMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],edsuMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],edsuMSPending_Emergency_ClientHighLessThan10.shape[0],edsuMSPending_Emergency_ClientHighGreaterThan10.shape[0],edsuMSPending_Emergency_ClientHighGreaterThan50.shape[0],edsuMSPending_Emergency_ClientHighGreaterThan100.shape[0],edsuMSPending_Emergency_ClientAverageLessThan10.shape[0],edsuMSPending_Emergency_ClientAverageGreaterThan10.shape[0],edsuMSPending_Emergency_ClientAverageGreaterThan50.shape[0],edsuMSPending_Emergency_ClientAverageGreaterThan100.shape[0],edsuMSPending_Emergency_ClientLowLessThan10.shape[0],edsuMSPending_Emergency_ClientLowGreaterThan10.shape[0],edsuMSPending_Emergency_ClientLowGreaterThan50.shape[0],edsuMSPending_Emergency_ClientLowGreaterThan100.shape[0],edsuMSPending_VendorCriticalLessThan10.shape[0],edsuMSPending_VendorCriticalGreaterThan10.shape[0],edsuMSPending_VendorCriticalGreaterThan50.shape[0],edsuMSPending_VendorCriticalGreaterThan100.shape[0],edsuMSPending_VendorHighLessThan10.shape[0],edsuMSPending_VendorHighGreaterThan10.shape[0],edsuMSPending_VendorHighGreaterThan50.shape[0],edsuMSPending_VendorHighGreaterThan100.shape[0],edsuMSPending_VendorAverageLessThan10.shape[0],edsuMSPending_VendorAverageGreaterThan10.shape[0],edsuMSPending_VendorAverageGreaterThan50.shape[0],edsuMSPending_VendorAverageGreaterThan100.shape[0],edsuMSPending_VendorLowLessThan10.shape[0],edsuMSPending_VendorLowGreaterThan10.shape[0],edsuMSPending_VendorLowGreaterThan50.shape[0],edsuMSPending_VendorLowGreaterThan100.shape[0],edsuMSWork_In_ProgressCriticalLessThan10.shape[0],edsuMSWork_In_ProgressCriticalGreaterThan10.shape[0],edsuMSWork_In_ProgressCriticalGreaterThan50.shape[0],edsuMSWork_In_ProgressCriticalGreaterThan100.shape[0],edsuMSWork_In_ProgressHighLessThan10.shape[0],edsuMSWork_In_ProgressHighGreaterThan10.shape[0],edsuMSWork_In_ProgressHighGreaterThan50.shape[0],edsuMSWork_In_ProgressHighGreaterThan100.shape[0],edsuMSWork_In_ProgressAverageLessThan10.shape[0],edsuMSWork_In_ProgressAverageGreaterThan10.shape[0],edsuMSWork_In_ProgressAverageGreaterThan50.shape[0],edsuMSWork_In_ProgressAverageGreaterThan100.shape[0],edsuMSWork_In_ProgressLowLessThan10.shape[0],edsuMSWork_In_ProgressLowGreaterThan10.shape[0],edsuMSWork_In_ProgressLowGreaterThan50.shape[0],edsuMSWork_In_ProgressLowGreaterThan100.shape[0],edsuNonMSOpenCriticalLessThan10.shape[0],edsuNonMSOpenCriticalGreaterThan10.shape[0],edsuNonMSOpenCriticalGreaterThan50.shape[0],edsuNonMSOpenCriticalGreaterThan100.shape[0],edsuNonMSOpenHighLessThan10.shape[0],edsuNonMSOpenHighGreaterThan10.shape[0],edsuNonMSOpenHighGreaterThan50.shape[0],edsuNonMSOpenHighGreaterThan100.shape[0],edsuNonMSOpenAverageLessThan10.shape[0],edsuNonMSOpenAverageGreaterThan10.shape[0],edsuNonMSOpenAverageGreaterThan50.shape[0],edsuNonMSOpenAverageGreaterThan100.shape[0],edsuNonMSOpenLowLessThan10.shape[0],edsuNonMSOpenLowGreaterThan10.shape[0],edsuNonMSOpenLowGreaterThan50.shape[0],edsuNonMSOpenLowGreaterThan100.shape[0],edsuNonMSPending_ClientCriticalLessThan10.shape[0],edsuNonMSPending_ClientCriticalGreaterThan10.shape[0],edsuNonMSPending_ClientCriticalGreaterThan50.shape[0],edsuNonMSPending_ClientCriticalGreaterThan100.shape[0],edsuNonMSPending_ClientHighLessThan10.shape[0],edsuNonMSPending_ClientHighGreaterThan10.shape[0],edsuNonMSPending_ClientHighGreaterThan50.shape[0],edsuNonMSPending_ClientHighGreaterThan100.shape[0],edsuNonMSPending_ClientAverageLessThan10.shape[0],edsuNonMSPending_ClientAverageGreaterThan10.shape[0],edsuNonMSPending_ClientAverageGreaterThan50.shape[0],edsuNonMSPending_ClientAverageGreaterThan100.shape[0],edsuNonMSPending_ClientLowLessThan10.shape[0],edsuNonMSPending_ClientLowGreaterThan10.shape[0],edsuNonMSPending_ClientLowGreaterThan50.shape[0],edsuNonMSPending_ClientLowGreaterThan100.shape[0],edsuNonMSPending_Emergency_ClientCriticalLessThan10.shape[0],edsuNonMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],edsuNonMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],edsuNonMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],edsuNonMSPending_Emergency_ClientHighLessThan10.shape[0],edsuNonMSPending_Emergency_ClientHighGreaterThan10.shape[0],edsuNonMSPending_Emergency_ClientHighGreaterThan50.shape[0],edsuNonMSPending_Emergency_ClientHighGreaterThan100.shape[0],edsuNonMSPending_Emergency_ClientAverageLessThan10.shape[0],edsuNonMSPending_Emergency_ClientAverageGreaterThan10.shape[0],edsuNonMSPending_Emergency_ClientAverageGreaterThan50.shape[0],edsuNonMSPending_Emergency_ClientAverageGreaterThan100.shape[0],edsuNonMSPending_Emergency_ClientLowLessThan10.shape[0],edsuNonMSPending_Emergency_ClientLowGreaterThan10.shape[0],edsuNonMSPending_Emergency_ClientLowGreaterThan50.shape[0],edsuNonMSPending_Emergency_ClientLowGreaterThan100.shape[0],edsuNonMSPending_VendorCriticalLessThan10.shape[0],edsuNonMSPending_VendorCriticalGreaterThan10.shape[0],edsuNonMSPending_VendorCriticalGreaterThan50.shape[0],edsuNonMSPending_VendorCriticalGreaterThan100.shape[0],edsuNonMSPending_VendorHighLessThan10.shape[0],edsuNonMSPending_VendorHighGreaterThan10.shape[0],edsuNonMSPending_VendorHighGreaterThan50.shape[0],edsuNonMSPending_VendorHighGreaterThan100.shape[0],edsuNonMSPending_VendorAverageLessThan10.shape[0],edsuNonMSPending_VendorAverageGreaterThan10.shape[0],edsuNonMSPending_VendorAverageGreaterThan50.shape[0],edsuNonMSPending_VendorAverageGreaterThan100.shape[0],edsuNonMSPending_VendorLowLessThan10.shape[0],edsuNonMSPending_VendorLowGreaterThan10.shape[0],edsuNonMSPending_VendorLowGreaterThan50.shape[0],edsuNonMSPending_VendorLowGreaterThan100.shape[0],edsuNonMSWork_In_ProgressCriticalLessThan10.shape[0],edsuNonMSWork_In_ProgressCriticalGreaterThan10.shape[0],edsuNonMSWork_In_ProgressCriticalGreaterThan50.shape[0],edsuNonMSWork_In_ProgressCriticalGreaterThan100.shape[0],edsuNonMSWork_In_ProgressHighLessThan10.shape[0],edsuNonMSWork_In_ProgressHighGreaterThan10.shape[0],edsuNonMSWork_In_ProgressHighGreaterThan50.shape[0],edsuNonMSWork_In_ProgressHighGreaterThan100.shape[0],edsuNonMSWork_In_ProgressAverageLessThan10.shape[0],edsuNonMSWork_In_ProgressAverageGreaterThan10.shape[0],edsuNonMSWork_In_ProgressAverageGreaterThan50.shape[0],edsuNonMSWork_In_ProgressAverageGreaterThan100.shape[0],edsuNonMSWork_In_ProgressLowLessThan10.shape[0],edsuNonMSWork_In_ProgressLowGreaterThan10.shape[0],edsuNonMSWork_In_ProgressLowGreaterThan50.shape[0],edsuNonMSWork_In_ProgressLowGreaterThan100.shape[0]
]}
rcdf=pd.DataFrame(rowcount)
rcdf.to_excel(rowpath,index=False)
print("success")

#rowcount
#edsuMSOpenCriticalLessThan10rc=edsuMSOpenCriticalLessThan10.shape[0]
##(edsuMSWork_In_ProgressHighLessThan10)
##(Msedsu.shape[0])
##(edsuMSOpen)
##(edsuMSOpenCritical)













#writer = pd.ExcelWriter(excelPath.replace(".xlsx","1.xlsx"), engine='xlsxwriter')

#edsu1.to_excel(writer, sheet_name=ng,index=False)
#edsupport1.to_excel(writer, sheet_name=es,index=False)
#ediAnalysts1.to_excel(writer, sheet_name=ea,index=False)
#webMethods1.to_excel(writer, sheet_name=we,index=False)

#writer.save()
##("success")
