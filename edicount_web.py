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
web=pd.read_excel(excelPath,sheet_name=0,index=False,engine='openpyxl')

#MS_NONMS
Msweb=web[web['MS_NONMS'] == "MS"]
NonMsweb=web[web['MS_NONMS'] == "Non_ms"]
#MS Status
webMSOpen=Msweb[Msweb['Status'] == "Open"]
webMSPending_Client=Msweb[Msweb['Status'] == "Pending Client"]
webMSPending_Emergency_Client=Msweb[Msweb['Status'] == "Pending Emergency Client"]
webMSPending_Vendor=Msweb[Msweb['Status'] == "Pending Vendor"]
webMSWork_In_Progress=Msweb[Msweb['Status'] == "Work In Progress"]
#(webMSWork_In_Progress)
#Non_MS Status
webNonMSOpen=NonMsweb[NonMsweb['Status'] == "Open"]
webNonMSPending_Client=NonMsweb[NonMsweb['Status'] == "Pending Client"]
webNonMSPending_Emergency_Client=NonMsweb[NonMsweb['Status'] == "Pending Emergency Client"]
webNonMSPending_Vendor=NonMsweb[NonMsweb['Status'] == "Pending Vendor"]
webNonMSWork_In_Progress=NonMsweb[NonMsweb['Status'] == "Work In Progress"]

#MS_OPEN_PRIORITY**
webMSOpenCritical=webMSOpen[webMSOpen['Priority'] == 1]
webMSOpenHigh=webMSOpen[webMSOpen['Priority'] == 2]
webMSOpenAverage=webMSOpen[webMSOpen['Priority'] == 3]
webMSOpenLow=webMSOpen[webMSOpen['Priority'] == 4]

##MS_Pending_Client_PRIORITY
webMSPending_ClientCritical=webMSPending_Client[webMSPending_Client['Priority'] == 1]
webMSPending_ClientHigh=webMSPending_Client[webMSPending_Client['Priority'] == 2]
webMSPending_ClientAverage=webMSPending_Client[webMSPending_Client['Priority'] == 3]
webMSPending_ClientLow=webMSPending_Client[webMSPending_Client['Priority'] == 4]

#MS_Pending_Emergency_Client
webMSPending_Emergency_ClientCritical=webMSPending_Emergency_Client[webMSPending_Emergency_Client['Priority'] == 1]
webMSPending_Emergency_ClientHigh=webMSPending_Emergency_Client[webMSPending_Emergency_Client['Priority'] == 2]
webMSPending_Emergency_ClientAverage=webMSPending_Emergency_Client[webMSPending_Emergency_Client['Priority'] == 3]
webMSPending_Emergency_ClientLow=webMSPending_Emergency_Client[webMSPending_Emergency_Client['Priority'] == 4]

#MS_Pending_Vendor
webMSPending_VendorCritical=webMSPending_Vendor[webMSPending_Vendor['Priority'] == 1]
webMSPending_VendorHigh=webMSPending_Vendor[webMSPending_Vendor['Priority'] == 2]
webMSPending_VendorAverage=webMSPending_Vendor[webMSPending_Vendor['Priority'] == 3]
webMSPending_VendorLow=webMSPending_Vendor[webMSPending_Vendor['Priority'] == 4]

#MS_Work_In_Progress
webMSWork_In_ProgressCritical=webMSWork_In_Progress[webMSWork_In_Progress['Priority'] == 1]
webMSWork_In_ProgressHigh=webMSWork_In_Progress[webMSWork_In_Progress['Priority'] == 2]
webMSWork_In_ProgressAverage=webMSWork_In_Progress[webMSWork_In_Progress['Priority'] == 3]
webMSWork_In_ProgressLow=webMSWork_In_Progress[webMSWork_In_Progress['Priority'] == 4]
#(webMSWork_In_ProgressHigh)


#NonMS_OPEN_PRIORITY
webNonMSOpenCritical=webNonMSOpen[webNonMSOpen['Priority'] == 1]
webNonMSOpenHigh=webNonMSOpen[webNonMSOpen['Priority'] == 2]
webNonMSOpenAverage=webNonMSOpen[webNonMSOpen['Priority'] == 3]
webNonMSOpenLow=webNonMSOpen[webNonMSOpen['Priority'] == 4]

##NonMS_Pending_Client_PRIORITY
webNonMSPending_ClientCritical=webNonMSPending_Client[webNonMSPending_Client['Priority'] == 1]
webNonMSPending_ClientHigh=webNonMSPending_Client[webNonMSPending_Client['Priority'] == 2]
webNonMSPending_ClientAverage=webNonMSPending_Client[webNonMSPending_Client['Priority'] == 3]
webNonMSPending_ClientLow=webNonMSPending_Client[webNonMSPending_Client['Priority'] == 4]

#NonMS_Pending_Emergency_Client
webNonMSPending_Emergency_ClientCritical=webNonMSPending_Emergency_Client[webNonMSPending_Emergency_Client['Priority'] == 1]
webNonMSPending_Emergency_ClientHigh=webNonMSPending_Emergency_Client[webNonMSPending_Emergency_Client['Priority'] == 2]
webNonMSPending_Emergency_ClientAverage=webNonMSPending_Emergency_Client[webNonMSPending_Emergency_Client['Priority'] == 3]
webNonMSPending_Emergency_ClientLow=webNonMSPending_Emergency_Client[webNonMSPending_Emergency_Client['Priority'] == 4]

#NonMS_Pending_Vendor
webNonMSPending_VendorCritical=webNonMSPending_Vendor[webNonMSPending_Vendor['Priority'] == 1]
webNonMSPending_VendorHigh=webNonMSPending_Vendor[webNonMSPending_Vendor['Priority'] == 2]
webNonMSPending_VendorAverage=webNonMSPending_Vendor[webNonMSPending_Vendor['Priority'] == 3]
webNonMSPending_VendorLow=webNonMSPending_Vendor[webNonMSPending_Vendor['Priority'] == 4]

#NonMS_Work_In_Progress
webNonMSWork_In_ProgressCritical=webNonMSWork_In_Progress[webNonMSWork_In_Progress['Priority'] == 1]
webNonMSWork_In_ProgressHigh=webNonMSWork_In_Progress[webNonMSWork_In_Progress['Priority'] == 2]
webNonMSWork_In_ProgressAverage=webNonMSWork_In_Progress[webNonMSWork_In_Progress['Priority'] == 3]
webNonMSWork_In_ProgressLow=webNonMSWork_In_Progress[webNonMSWork_In_Progress['Priority'] == 4]

#webMSOpenCriticalAging
webMSOpenCriticalLessThan10=webMSOpenCritical[webMSOpenCritical['Ageing days'] <=10]
webMSOpenCriticalGreaterThan10=webMSOpenCritical[webMSOpenCritical['Ageing days'] > 10]
webMSOpenCriticalGreaterThan50=webMSOpenCritical[webMSOpenCritical['Ageing days'] > 50]
webMSOpenCriticalGreaterThan100=webMSOpenCritical[webMSOpenCritical['Ageing days'] > 100]

#webMSOpenHighAging
webMSOpenHighLessThan10=webMSOpenHigh[webMSOpenHigh['Ageing days'] <=10]
webMSOpenHighGreaterThan10=webMSOpenHigh[webMSOpenHigh['Ageing days'] > 10]
webMSOpenHighGreaterThan50=webMSOpenHigh[webMSOpenHigh['Ageing days'] > 50]
webMSOpenHighGreaterThan100=webMSOpenHigh[webMSOpenHigh['Ageing days'] > 100]

#webMSOpenAverageAging
webMSOpenAverageLessThan10=webMSOpenAverage[webMSOpenAverage['Ageing days'] <=10]
webMSOpenAverageGreaterThan10=webMSOpenAverage[webMSOpenAverage['Ageing days'] > 10]
webMSOpenAverageGreaterThan50=webMSOpenAverage[webMSOpenAverage['Ageing days'] > 50]
webMSOpenAverageGreaterThan100=webMSOpenAverage[webMSOpenAverage['Ageing days'] > 100]

#webMSOpenLowAging
webMSOpenLowLessThan10=webMSOpenLow[webMSOpenLow['Ageing days'] <=10]
webMSOpenLowGreaterThan10=webMSOpenLow[webMSOpenLow['Ageing days'] > 10]
webMSOpenLowGreaterThan50=webMSOpenLow[webMSOpenLow['Ageing days'] > 50]
webMSOpenLowGreaterThan100=webMSOpenLow[webMSOpenLow['Ageing days'] > 100]


#webMSPending_ClientCriticalAging
webMSPending_ClientCriticalLessThan10=webMSPending_ClientCritical[webMSPending_ClientCritical['Ageing days'] <=10]
webMSPending_ClientCriticalGreaterThan10=webMSPending_ClientCritical[webMSPending_ClientCritical['Ageing days'] > 10]
webMSPending_ClientCriticalGreaterThan50=webMSPending_ClientCritical[webMSPending_ClientCritical['Ageing days'] > 50]
webMSPending_ClientCriticalGreaterThan100=webMSPending_ClientCritical[webMSPending_ClientCritical['Ageing days'] > 100]

#webMSPending_ClientHighAging
webMSPending_ClientHighLessThan10=webMSPending_ClientHigh[webMSPending_ClientHigh['Ageing days'] <=10]
webMSPending_ClientHighGreaterThan10=webMSPending_ClientHigh[webMSPending_ClientHigh['Ageing days'] > 10]
webMSPending_ClientHighGreaterThan50=webMSPending_ClientHigh[webMSPending_ClientHigh['Ageing days'] > 50]
webMSPending_ClientHighGreaterThan100=webMSPending_ClientHigh[webMSPending_ClientHigh['Ageing days'] > 100]

#webMSPending_ClientAverageAging
webMSPending_ClientAverageLessThan10=webMSPending_ClientAverage[webMSPending_ClientAverage['Ageing days'] <=10]
webMSPending_ClientAverageGreaterThan10=webMSPending_ClientAverage[webMSPending_ClientAverage['Ageing days'] > 10]
webMSPending_ClientAverageGreaterThan50=webMSPending_ClientAverage[webMSPending_ClientAverage['Ageing days'] > 50]
webMSPending_ClientAverageGreaterThan100=webMSPending_ClientAverage[webMSPending_ClientAverage['Ageing days'] > 100]

#webMSPending_ClientLowAging
webMSPending_ClientLowLessThan10=webMSPending_ClientLow[webMSPending_ClientLow['Ageing days'] <=10]
webMSPending_ClientLowGreaterThan10=webMSPending_ClientLow[webMSPending_ClientLow['Ageing days'] > 10]
webMSPending_ClientLowGreaterThan50=webMSPending_ClientLow[webMSPending_ClientLow['Ageing days'] > 50]
webMSPending_ClientLowGreaterThan100=webMSPending_ClientLow[webMSPending_ClientLow['Ageing days'] > 100]


#webMSPending_Emergency_ClientCriticalAging
webMSPending_Emergency_ClientCriticalLessThan10=webMSPending_Emergency_ClientCritical[webMSPending_Emergency_ClientCritical['Ageing days'] <=10]
webMSPending_Emergency_ClientCriticalGreaterThan10=webMSPending_Emergency_ClientCritical[webMSPending_Emergency_ClientCritical['Ageing days'] > 10]
webMSPending_Emergency_ClientCriticalGreaterThan50=webMSPending_Emergency_ClientCritical[webMSPending_Emergency_ClientCritical['Ageing days'] > 50]
webMSPending_Emergency_ClientCriticalGreaterThan100=webMSPending_Emergency_ClientCritical[webMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#webMSPending_Emergency_ClientHighAging
webMSPending_Emergency_ClientHighLessThan10=webMSPending_Emergency_ClientHigh[webMSPending_Emergency_ClientHigh['Ageing days'] <=10]
webMSPending_Emergency_ClientHighGreaterThan10=webMSPending_Emergency_ClientHigh[webMSPending_Emergency_ClientHigh['Ageing days'] > 10]
webMSPending_Emergency_ClientHighGreaterThan50=webMSPending_Emergency_ClientHigh[webMSPending_Emergency_ClientHigh['Ageing days'] > 50]
webMSPending_Emergency_ClientHighGreaterThan100=webMSPending_Emergency_ClientHigh[webMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#webMSPending_Emergency_ClientAverageAging
webMSPending_Emergency_ClientAverageLessThan10=webMSPending_Emergency_ClientAverage[webMSPending_Emergency_ClientAverage['Ageing days'] <=10]
webMSPending_Emergency_ClientAverageGreaterThan10=webMSPending_Emergency_ClientAverage[webMSPending_Emergency_ClientAverage['Ageing days'] > 10]
webMSPending_Emergency_ClientAverageGreaterThan50=webMSPending_Emergency_ClientAverage[webMSPending_Emergency_ClientAverage['Ageing days'] > 50]
webMSPending_Emergency_ClientAverageGreaterThan100=webMSPending_Emergency_ClientAverage[webMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#webMSPending_Emergency_ClientLowAging
webMSPending_Emergency_ClientLowLessThan10=webMSPending_Emergency_ClientLow[webMSPending_Emergency_ClientLow['Ageing days'] <=10]
webMSPending_Emergency_ClientLowGreaterThan10=webMSPending_Emergency_ClientLow[webMSPending_Emergency_ClientLow['Ageing days'] > 10]
webMSPending_Emergency_ClientLowGreaterThan50=webMSPending_Emergency_ClientLow[webMSPending_Emergency_ClientLow['Ageing days'] > 50]
webMSPending_Emergency_ClientLowGreaterThan100=webMSPending_Emergency_ClientLow[webMSPending_Emergency_ClientLow['Ageing days'] > 100]


#webMSPending_VendorCriticalAging
webMSPending_VendorCriticalLessThan10=webMSPending_VendorCritical[webMSPending_VendorCritical['Ageing days'] <=10]
webMSPending_VendorCriticalGreaterThan10=webMSPending_VendorCritical[webMSPending_VendorCritical['Ageing days'] > 10]
webMSPending_VendorCriticalGreaterThan50=webMSPending_VendorCritical[webMSPending_VendorCritical['Ageing days'] > 50]
webMSPending_VendorCriticalGreaterThan100=webMSPending_VendorCritical[webMSPending_VendorCritical['Ageing days'] > 100]

#webMSPending_VendorHighAging
webMSPending_VendorHighLessThan10=webMSPending_VendorHigh[webMSPending_VendorHigh['Ageing days'] <=10]
webMSPending_VendorHighGreaterThan10=webMSPending_VendorHigh[webMSPending_VendorHigh['Ageing days'] > 10]
webMSPending_VendorHighGreaterThan50=webMSPending_VendorHigh[webMSPending_VendorHigh['Ageing days'] > 50]
webMSPending_VendorHighGreaterThan100=webMSPending_VendorHigh[webMSPending_VendorHigh['Ageing days'] > 100]

#webMSPending_VendorAverageAging
webMSPending_VendorAverageLessThan10=webMSPending_VendorAverage[webMSPending_VendorAverage['Ageing days'] <=10]
webMSPending_VendorAverageGreaterThan10=webMSPending_VendorAverage[webMSPending_VendorAverage['Ageing days'] > 10]
webMSPending_VendorAverageGreaterThan50=webMSPending_VendorAverage[webMSPending_VendorAverage['Ageing days'] > 50]
webMSPending_VendorAverageGreaterThan100=webMSPending_VendorAverage[webMSPending_VendorAverage['Ageing days'] > 100]

#webMSPending_VendorLowAging
webMSPending_VendorLowLessThan10=webMSPending_VendorLow[webMSPending_VendorLow['Ageing days'] <=10]
webMSPending_VendorLowGreaterThan10=webMSPending_VendorLow[webMSPending_VendorLow['Ageing days'] > 10]
webMSPending_VendorLowGreaterThan50=webMSPending_VendorLow[webMSPending_VendorLow['Ageing days'] > 50]
webMSPending_VendorLowGreaterThan100=webMSPending_VendorLow[webMSPending_VendorLow['Ageing days'] > 100]


#webMSWork_In_ProgressCriticalAging
webMSWork_In_ProgressCriticalLessThan10=webMSWork_In_ProgressCritical[webMSWork_In_ProgressCritical['Ageing days'] <=10]
webMSWork_In_ProgressCriticalGreaterThan10=webMSWork_In_ProgressCritical[webMSWork_In_ProgressCritical['Ageing days'] > 10]
webMSWork_In_ProgressCriticalGreaterThan50=webMSWork_In_ProgressCritical[webMSWork_In_ProgressCritical['Ageing days'] > 50]
webMSWork_In_ProgressCriticalGreaterThan100=webMSWork_In_ProgressCritical[webMSWork_In_ProgressCritical['Ageing days'] > 100]

#webMSWork_In_ProgressHighAging
webMSWork_In_ProgressHighLessThan10=webMSWork_In_ProgressHigh[webMSWork_In_ProgressHigh['Ageing days'] <=10]
webMSWork_In_ProgressHighGreaterThan10=webMSWork_In_ProgressHigh[webMSWork_In_ProgressHigh['Ageing days'] > 10]
webMSWork_In_ProgressHighGreaterThan50=webMSWork_In_ProgressHigh[webMSWork_In_ProgressHigh['Ageing days'] > 50]
webMSWork_In_ProgressHighGreaterThan100=webMSWork_In_ProgressHigh[webMSWork_In_ProgressHigh['Ageing days'] > 100]
#(webMSWork_In_ProgressHighLessThan10)
#(webMSWork_In_ProgressHighLessThan10.shape[0])
#webMSWork_In_ProgressAverageAging
webMSWork_In_ProgressAverageLessThan10=webMSWork_In_ProgressAverage[webMSWork_In_ProgressAverage['Ageing days'] <=10]
webMSWork_In_ProgressAverageGreaterThan10=webMSWork_In_ProgressAverage[webMSWork_In_ProgressAverage['Ageing days'] > 10]
webMSWork_In_ProgressAverageGreaterThan50=webMSWork_In_ProgressAverage[webMSWork_In_ProgressAverage['Ageing days'] > 50]
webMSWork_In_ProgressAverageGreaterThan100=webMSWork_In_ProgressAverage[webMSWork_In_ProgressAverage['Ageing days'] > 100]

#webMSWork_In_ProgressLowAging
webMSWork_In_ProgressLowLessThan10=webMSWork_In_ProgressLow[webMSWork_In_ProgressLow['Ageing days'] <=10]
webMSWork_In_ProgressLowGreaterThan10=webMSWork_In_ProgressLow[webMSWork_In_ProgressLow['Ageing days'] > 10]
webMSWork_In_ProgressLowGreaterThan50=webMSWork_In_ProgressLow[webMSWork_In_ProgressLow['Ageing days'] > 50]
webMSWork_In_ProgressLowGreaterThan100=webMSWork_In_ProgressLow[webMSWork_In_ProgressLow['Ageing days'] > 100]



#webNonMSOpenCriticalAging
webNonMSOpenCriticalLessThan10=webNonMSOpenCritical[webNonMSOpenCritical['Ageing days'] <=10]
webNonMSOpenCriticalGreaterThan10=webNonMSOpenCritical[webNonMSOpenCritical['Ageing days'] > 10]
webNonMSOpenCriticalGreaterThan50=webNonMSOpenCritical[webNonMSOpenCritical['Ageing days'] > 50]
webNonMSOpenCriticalGreaterThan100=webNonMSOpenCritical[webNonMSOpenCritical['Ageing days'] > 100]

#webNonMSOpenHighAging
webNonMSOpenHighLessThan10=webNonMSOpenHigh[webNonMSOpenHigh['Ageing days'] <=10]
webNonMSOpenHighGreaterThan10=webNonMSOpenHigh[webNonMSOpenHigh['Ageing days'] > 10]
webNonMSOpenHighGreaterThan50=webNonMSOpenHigh[webNonMSOpenHigh['Ageing days'] > 50]
webNonMSOpenHighGreaterThan100=webNonMSOpenHigh[webNonMSOpenHigh['Ageing days'] > 100]

#webNonMSOpenAverageAging
webNonMSOpenAverageLessThan10=webNonMSOpenAverage[webNonMSOpenAverage['Ageing days'] <=10]
webNonMSOpenAverageGreaterThan10=webNonMSOpenAverage[webNonMSOpenAverage['Ageing days'] > 10]
webNonMSOpenAverageGreaterThan50=webNonMSOpenAverage[webNonMSOpenAverage['Ageing days'] > 50]
webNonMSOpenAverageGreaterThan100=webNonMSOpenAverage[webNonMSOpenAverage['Ageing days'] > 100]

#webNonMSOpenLowAging
webNonMSOpenLowLessThan10=webNonMSOpenLow[webNonMSOpenLow['Ageing days'] <=10]
webNonMSOpenLowGreaterThan10=webNonMSOpenLow[webNonMSOpenLow['Ageing days'] > 10]
webNonMSOpenLowGreaterThan50=webNonMSOpenLow[webNonMSOpenLow['Ageing days'] > 50]
webNonMSOpenLowGreaterThan100=webNonMSOpenLow[webNonMSOpenLow['Ageing days'] > 100]


#webNonMSPending_ClientCriticalAging
webNonMSPending_ClientCriticalLessThan10=webNonMSPending_ClientCritical[webNonMSPending_ClientCritical['Ageing days'] <=10]
webNonMSPending_ClientCriticalGreaterThan10=webNonMSPending_ClientCritical[webNonMSPending_ClientCritical['Ageing days'] > 10]
webNonMSPending_ClientCriticalGreaterThan50=webNonMSPending_ClientCritical[webNonMSPending_ClientCritical['Ageing days'] > 50]
webNonMSPending_ClientCriticalGreaterThan100=webNonMSPending_ClientCritical[webNonMSPending_ClientCritical['Ageing days'] > 100]

#webNonMSPending_ClientHighAging
webNonMSPending_ClientHighLessThan10=webNonMSPending_ClientHigh[webNonMSPending_ClientHigh['Ageing days'] <=10]
webNonMSPending_ClientHighGreaterThan10=webNonMSPending_ClientHigh[webNonMSPending_ClientHigh['Ageing days'] > 10]
webNonMSPending_ClientHighGreaterThan50=webNonMSPending_ClientHigh[webNonMSPending_ClientHigh['Ageing days'] > 50]
webNonMSPending_ClientHighGreaterThan100=webNonMSPending_ClientHigh[webNonMSPending_ClientHigh['Ageing days'] > 100]

#webNonMSPending_ClientAverageAging
webNonMSPending_ClientAverageLessThan10=webNonMSPending_ClientAverage[webNonMSPending_ClientAverage['Ageing days'] <=10]
webNonMSPending_ClientAverageGreaterThan10=webNonMSPending_ClientAverage[webNonMSPending_ClientAverage['Ageing days'] > 10]
webNonMSPending_ClientAverageGreaterThan50=webNonMSPending_ClientAverage[webNonMSPending_ClientAverage['Ageing days'] > 50]
webNonMSPending_ClientAverageGreaterThan100=webNonMSPending_ClientAverage[webNonMSPending_ClientAverage['Ageing days'] > 100]

#webNonMSPending_ClientLowAging
webNonMSPending_ClientLowLessThan10=webNonMSPending_ClientLow[webNonMSPending_ClientLow['Ageing days'] <=10]
webNonMSPending_ClientLowGreaterThan10=webNonMSPending_ClientLow[webNonMSPending_ClientLow['Ageing days'] > 10]
webNonMSPending_ClientLowGreaterThan50=webNonMSPending_ClientLow[webNonMSPending_ClientLow['Ageing days'] > 50]
webNonMSPending_ClientLowGreaterThan100=webNonMSPending_ClientLow[webNonMSPending_ClientLow['Ageing days'] > 100]


#webNonMSPending_Emergency_ClientCriticalAging
webNonMSPending_Emergency_ClientCriticalLessThan10=webNonMSPending_Emergency_ClientCritical[webNonMSPending_Emergency_ClientCritical['Ageing days'] <=10]
webNonMSPending_Emergency_ClientCriticalGreaterThan10=webNonMSPending_Emergency_ClientCritical[webNonMSPending_Emergency_ClientCritical['Ageing days'] > 10]
webNonMSPending_Emergency_ClientCriticalGreaterThan50=webNonMSPending_Emergency_ClientCritical[webNonMSPending_Emergency_ClientCritical['Ageing days'] > 50]
webNonMSPending_Emergency_ClientCriticalGreaterThan100=webNonMSPending_Emergency_ClientCritical[webNonMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#webNonMSPending_Emergency_ClientHighAging
webNonMSPending_Emergency_ClientHighLessThan10=webNonMSPending_Emergency_ClientHigh[webNonMSPending_Emergency_ClientHigh['Ageing days'] <=10]
webNonMSPending_Emergency_ClientHighGreaterThan10=webNonMSPending_Emergency_ClientHigh[webNonMSPending_Emergency_ClientHigh['Ageing days'] > 10]
webNonMSPending_Emergency_ClientHighGreaterThan50=webNonMSPending_Emergency_ClientHigh[webNonMSPending_Emergency_ClientHigh['Ageing days'] > 50]
webNonMSPending_Emergency_ClientHighGreaterThan100=webNonMSPending_Emergency_ClientHigh[webNonMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#webNonMSPending_Emergency_ClientAverageAging
webNonMSPending_Emergency_ClientAverageLessThan10=webNonMSPending_Emergency_ClientAverage[webNonMSPending_Emergency_ClientAverage['Ageing days'] <=10]
webNonMSPending_Emergency_ClientAverageGreaterThan10=webNonMSPending_Emergency_ClientAverage[webNonMSPending_Emergency_ClientAverage['Ageing days'] > 10]
webNonMSPending_Emergency_ClientAverageGreaterThan50=webNonMSPending_Emergency_ClientAverage[webNonMSPending_Emergency_ClientAverage['Ageing days'] > 50]
webNonMSPending_Emergency_ClientAverageGreaterThan100=webNonMSPending_Emergency_ClientAverage[webNonMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#webNonMSPending_Emergency_ClientLowAging
webNonMSPending_Emergency_ClientLowLessThan10=webNonMSPending_Emergency_ClientLow[webNonMSPending_Emergency_ClientLow['Ageing days'] <=10]
webNonMSPending_Emergency_ClientLowGreaterThan10=webNonMSPending_Emergency_ClientLow[webNonMSPending_Emergency_ClientLow['Ageing days'] > 10]
webNonMSPending_Emergency_ClientLowGreaterThan50=webNonMSPending_Emergency_ClientLow[webNonMSPending_Emergency_ClientLow['Ageing days'] > 50]
webNonMSPending_Emergency_ClientLowGreaterThan100=webNonMSPending_Emergency_ClientLow[webNonMSPending_Emergency_ClientLow['Ageing days'] > 100]


#webNonMSPending_VendorCriticalAging
webNonMSPending_VendorCriticalLessThan10=webNonMSPending_VendorCritical[webNonMSPending_VendorCritical['Ageing days'] <=10]
webNonMSPending_VendorCriticalGreaterThan10=webNonMSPending_VendorCritical[webNonMSPending_VendorCritical['Ageing days'] > 10]
webNonMSPending_VendorCriticalGreaterThan50=webNonMSPending_VendorCritical[webNonMSPending_VendorCritical['Ageing days'] > 50]
webNonMSPending_VendorCriticalGreaterThan100=webNonMSPending_VendorCritical[webNonMSPending_VendorCritical['Ageing days'] > 100]

#webNonMSPending_VendorHighAging
webNonMSPending_VendorHighLessThan10=webNonMSPending_VendorHigh[webNonMSPending_VendorHigh['Ageing days'] <=10]
webNonMSPending_VendorHighGreaterThan10=webNonMSPending_VendorHigh[webNonMSPending_VendorHigh['Ageing days'] > 10]
webNonMSPending_VendorHighGreaterThan50=webNonMSPending_VendorHigh[webNonMSPending_VendorHigh['Ageing days'] > 50]
webNonMSPending_VendorHighGreaterThan100=webNonMSPending_VendorHigh[webNonMSPending_VendorHigh['Ageing days'] > 100]

#webNonMSPending_VendorAverageAging
webNonMSPending_VendorAverageLessThan10=webNonMSPending_VendorAverage[webNonMSPending_VendorAverage['Ageing days'] <=10]
webNonMSPending_VendorAverageGreaterThan10=webNonMSPending_VendorAverage[webNonMSPending_VendorAverage['Ageing days'] > 10]
webNonMSPending_VendorAverageGreaterThan50=webNonMSPending_VendorAverage[webNonMSPending_VendorAverage['Ageing days'] > 50]
webNonMSPending_VendorAverageGreaterThan100=webNonMSPending_VendorAverage[webNonMSPending_VendorAverage['Ageing days'] > 100]

#webNonMSPending_VendorLowAging
webNonMSPending_VendorLowLessThan10=webNonMSPending_VendorLow[webNonMSPending_VendorLow['Ageing days'] <=10]
webNonMSPending_VendorLowGreaterThan10=webNonMSPending_VendorLow[webNonMSPending_VendorLow['Ageing days'] > 10]
webNonMSPending_VendorLowGreaterThan50=webNonMSPending_VendorLow[webNonMSPending_VendorLow['Ageing days'] > 50]
webNonMSPending_VendorLowGreaterThan100=webNonMSPending_VendorLow[webNonMSPending_VendorLow['Ageing days'] > 100]


#webNonMSWork_In_ProgressCriticalAging
webNonMSWork_In_ProgressCriticalLessThan10=webNonMSWork_In_ProgressCritical[webNonMSWork_In_ProgressCritical['Ageing days'] <=10]
webNonMSWork_In_ProgressCriticalGreaterThan10=webNonMSWork_In_ProgressCritical[webNonMSWork_In_ProgressCritical['Ageing days'] > 10]
webNonMSWork_In_ProgressCriticalGreaterThan50=webNonMSWork_In_ProgressCritical[webNonMSWork_In_ProgressCritical['Ageing days'] > 50]
webNonMSWork_In_ProgressCriticalGreaterThan100=webNonMSWork_In_ProgressCritical[webNonMSWork_In_ProgressCritical['Ageing days'] > 100]

#webNonMSWork_In_ProgressHighAging
webNonMSWork_In_ProgressHighLessThan10=webNonMSWork_In_ProgressHigh[webNonMSWork_In_ProgressHigh['Ageing days'] <=10]
webNonMSWork_In_ProgressHighGreaterThan10=webNonMSWork_In_ProgressHigh[webNonMSWork_In_ProgressHigh['Ageing days'] > 10]
webNonMSWork_In_ProgressHighGreaterThan50=webNonMSWork_In_ProgressHigh[webNonMSWork_In_ProgressHigh['Ageing days'] > 50]
webNonMSWork_In_ProgressHighGreaterThan100=webNonMSWork_In_ProgressHigh[webNonMSWork_In_ProgressHigh['Ageing days'] > 100]

#webNonMSWork_In_ProgressAverageAging
webNonMSWork_In_ProgressAverageLessThan10=webNonMSWork_In_ProgressAverage[webNonMSWork_In_ProgressAverage['Ageing days'] <=10]
webNonMSWork_In_ProgressAverageGreaterThan10=webNonMSWork_In_ProgressAverage[webNonMSWork_In_ProgressAverage['Ageing days'] > 10]
webNonMSWork_In_ProgressAverageGreaterThan50=webNonMSWork_In_ProgressAverage[webNonMSWork_In_ProgressAverage['Ageing days'] > 50]
webNonMSWork_In_ProgressAverageGreaterThan100=webNonMSWork_In_ProgressAverage[webNonMSWork_In_ProgressAverage['Ageing days'] > 100]

#webNonMSWork_In_ProgressLowAging
webNonMSWork_In_ProgressLowLessThan10=webNonMSWork_In_ProgressLow[webNonMSWork_In_ProgressLow['Ageing days'] <=10]
webNonMSWork_In_ProgressLowGreaterThan10=webNonMSWork_In_ProgressLow[webNonMSWork_In_ProgressLow['Ageing days'] > 10]
webNonMSWork_In_ProgressLowGreaterThan50=webNonMSWork_In_ProgressLow[webNonMSWork_In_ProgressLow['Ageing days'] > 50]
webNonMSWork_In_ProgressLowGreaterThan100=webNonMSWork_In_ProgressLow[webNonMSWork_In_ProgressLow['Ageing days'] > 100]


rowcount={'Data':['webMSOpenCriticalLessThan10','webMSOpenCriticalGreaterThan10','webMSOpenCriticalGreaterThan50','webMSOpenCriticalGreaterThan100','webMSOpenHighLessThan10','webMSOpenHighGreaterThan10','webMSOpenHighGreaterThan50','webMSOpenHighGreaterThan100','webMSOpenAverageLessThan10','webMSOpenAverageGreaterThan10','webMSOpenAverageGreaterThan50','webMSOpenAverageGreaterThan100','webMSOpenLowLessThan10','webMSOpenLowGreaterThan10','webMSOpenLowGreaterThan50','webMSOpenLowGreaterThan100','webMSPending_ClientCriticalLessThan10','webMSPending_ClientCriticalGreaterThan10','webMSPending_ClientCriticalGreaterThan50','webMSPending_ClientCriticalGreaterThan100','webMSPending_ClientHighLessThan10','webMSPending_ClientHighGreaterThan10','webMSPending_ClientHighGreaterThan50','webMSPending_ClientHighGreaterThan100','webMSPending_ClientAverageLessThan10','webMSPending_ClientAverageGreaterThan10','webMSPending_ClientAverageGreaterThan50','webMSPending_ClientAverageGreaterThan100','webMSPending_ClientLowLessThan10','webMSPending_ClientLowGreaterThan10','webMSPending_ClientLowGreaterThan50','webMSPending_ClientLowGreaterThan100','webMSPending_Emergency_ClientCriticalLessThan10','webMSPending_Emergency_ClientCriticalGreaterThan10','webMSPending_Emergency_ClientCriticalGreaterThan50','webMSPending_Emergency_ClientCriticalGreaterThan100','webMSPending_Emergency_ClientHighLessThan10','webMSPending_Emergency_ClientHighGreaterThan10','webMSPending_Emergency_ClientHighGreaterThan50','webMSPending_Emergency_ClientHighGreaterThan100','webMSPending_Emergency_ClientAverageLessThan10','webMSPending_Emergency_ClientAverageGreaterThan10','webMSPending_Emergency_ClientAverageGreaterThan50','webMSPending_Emergency_ClientAverageGreaterThan100','webMSPending_Emergency_ClientLowLessThan10','webMSPending_Emergency_ClientLowGreaterThan10','webMSPending_Emergency_ClientLowGreaterThan50','webMSPending_Emergency_ClientLowGreaterThan100','webMSPending_VendorCriticalLessThan10','webMSPending_VendorCriticalGreaterThan10','webMSPending_VendorCriticalGreaterThan50','webMSPending_VendorCriticalGreaterThan100','webMSPending_VendorHighLessThan10','webMSPending_VendorHighGreaterThan10','webMSPending_VendorHighGreaterThan50','webMSPending_VendorHighGreaterThan100','webMSPending_VendorAverageLessThan10','webMSPending_VendorAverageGreaterThan10','webMSPending_VendorAverageGreaterThan50','webMSPending_VendorAverageGreaterThan100','webMSPending_VendorLowLessThan10','webMSPending_VendorLowGreaterThan10','webMSPending_VendorLowGreaterThan50','webMSPending_VendorLowGreaterThan100','webMSWork_In_ProgressCriticalLessThan10','webMSWork_In_ProgressCriticalGreaterThan10','webMSWork_In_ProgressCriticalGreaterThan50','webMSWork_In_ProgressCriticalGreaterThan100','webMSWork_In_ProgressHighLessThan10','webMSWork_In_ProgressHighGreaterThan10','webMSWork_In_ProgressHighGreaterThan50','webMSWork_In_ProgressHighGreaterThan100','webMSWork_In_ProgressAverageLessThan10','webMSWork_In_ProgressAverageGreaterThan10','webMSWork_In_ProgressAverageGreaterThan50','webMSWork_In_ProgressAverageGreaterThan100','webMSWork_In_ProgressLowLessThan10','webMSWork_In_ProgressLowGreaterThan10','webMSWork_In_ProgressLowGreaterThan50','webMSWork_In_ProgressLowGreaterThan100','webNonMSOpenCriticalLessThan10','webNonMSOpenCriticalGreaterThan10','webNonMSOpenCriticalGreaterThan50','webNonMSOpenCriticalGreaterThan100','webNonMSOpenHighLessThan10','webNonMSOpenHighGreaterThan10','webNonMSOpenHighGreaterThan50','webNonMSOpenHighGreaterThan100','webNonMSOpenAverageLessThan10','webNonMSOpenAverageGreaterThan10','webNonMSOpenAverageGreaterThan50','webNonMSOpenAverageGreaterThan100','webNonMSOpenLowLessThan10','webNonMSOpenLowGreaterThan10','webNonMSOpenLowGreaterThan50','webNonMSOpenLowGreaterThan100','webNonMSPending_ClientCriticalLessThan10','webNonMSPending_ClientCriticalGreaterThan10','webNonMSPending_ClientCriticalGreaterThan50','webNonMSPending_ClientCriticalGreaterThan100','webNonMSPending_ClientHighLessThan10','webNonMSPending_ClientHighGreaterThan10','webNonMSPending_ClientHighGreaterThan50','webNonMSPending_ClientHighGreaterThan100','webNonMSPending_ClientAverageLessThan10','webNonMSPending_ClientAverageGreaterThan10','webNonMSPending_ClientAverageGreaterThan50','webNonMSPending_ClientAverageGreaterThan100','webNonMSPending_ClientLowLessThan10','webNonMSPending_ClientLowGreaterThan10','webNonMSPending_ClientLowGreaterThan50','webNonMSPending_ClientLowGreaterThan100','webNonMSPending_Emergency_ClientCriticalLessThan10','webNonMSPending_Emergency_ClientCriticalGreaterThan10','webNonMSPending_Emergency_ClientCriticalGreaterThan50','webNonMSPending_Emergency_ClientCriticalGreaterThan100','webNonMSPending_Emergency_ClientHighLessThan10','webNonMSPending_Emergency_ClientHighGreaterThan10','webNonMSPending_Emergency_ClientHighGreaterThan50','webNonMSPending_Emergency_ClientHighGreaterThan100','webNonMSPending_Emergency_ClientAverageLessThan10','webNonMSPending_Emergency_ClientAverageGreaterThan10','webNonMSPending_Emergency_ClientAverageGreaterThan50','webNonMSPending_Emergency_ClientAverageGreaterThan100','webNonMSPending_Emergency_ClientLowLessThan10','webNonMSPending_Emergency_ClientLowGreaterThan10','webNonMSPending_Emergency_ClientLowGreaterThan50','webNonMSPending_Emergency_ClientLowGreaterThan100','webNonMSPending_VendorCriticalLessThan10','webNonMSPending_VendorCriticalGreaterThan10','webNonMSPending_VendorCriticalGreaterThan50','webNonMSPending_VendorCriticalGreaterThan100','webNonMSPending_VendorHighLessThan10','webNonMSPending_VendorHighGreaterThan10','webNonMSPending_VendorHighGreaterThan50','webNonMSPending_VendorHighGreaterThan100','webNonMSPending_VendorAverageLessThan10','webNonMSPending_VendorAverageGreaterThan10','webNonMSPending_VendorAverageGreaterThan50','webNonMSPending_VendorAverageGreaterThan100','webNonMSPending_VendorLowLessThan10','webNonMSPending_VendorLowGreaterThan10','webNonMSPending_VendorLowGreaterThan50','webNonMSPending_VendorLowGreaterThan100','webNonMSWork_In_ProgressCriticalLessThan10','webNonMSWork_In_ProgressCriticalGreaterThan10','webNonMSWork_In_ProgressCriticalGreaterThan50','webNonMSWork_In_ProgressCriticalGreaterThan100','webNonMSWork_In_ProgressHighLessThan10','webNonMSWork_In_ProgressHighGreaterThan10','webNonMSWork_In_ProgressHighGreaterThan50','webNonMSWork_In_ProgressHighGreaterThan100','webNonMSWork_In_ProgressAverageLessThan10','webNonMSWork_In_ProgressAverageGreaterThan10','webNonMSWork_In_ProgressAverageGreaterThan50','webNonMSWork_In_ProgressAverageGreaterThan100','webNonMSWork_In_ProgressLowLessThan10','webNonMSWork_In_ProgressLowGreaterThan10','webNonMSWork_In_ProgressLowGreaterThan50','webNonMSWork_In_ProgressLowGreaterThan100'
],'RowCount':[webMSOpenCriticalLessThan10.shape[0],webMSOpenCriticalGreaterThan10.shape[0],webMSOpenCriticalGreaterThan50.shape[0],webMSOpenCriticalGreaterThan100.shape[0],webMSOpenHighLessThan10.shape[0],webMSOpenHighGreaterThan10.shape[0],webMSOpenHighGreaterThan50.shape[0],webMSOpenHighGreaterThan100.shape[0],webMSOpenAverageLessThan10.shape[0],webMSOpenAverageGreaterThan10.shape[0],webMSOpenAverageGreaterThan50.shape[0],webMSOpenAverageGreaterThan100.shape[0],webMSOpenLowLessThan10.shape[0],webMSOpenLowGreaterThan10.shape[0],webMSOpenLowGreaterThan50.shape[0],webMSOpenLowGreaterThan100.shape[0],webMSPending_ClientCriticalLessThan10.shape[0],webMSPending_ClientCriticalGreaterThan10.shape[0],webMSPending_ClientCriticalGreaterThan50.shape[0],webMSPending_ClientCriticalGreaterThan100.shape[0],webMSPending_ClientHighLessThan10.shape[0],webMSPending_ClientHighGreaterThan10.shape[0],webMSPending_ClientHighGreaterThan50.shape[0],webMSPending_ClientHighGreaterThan100.shape[0],webMSPending_ClientAverageLessThan10.shape[0],webMSPending_ClientAverageGreaterThan10.shape[0],webMSPending_ClientAverageGreaterThan50.shape[0],webMSPending_ClientAverageGreaterThan100.shape[0],webMSPending_ClientLowLessThan10.shape[0],webMSPending_ClientLowGreaterThan10.shape[0],webMSPending_ClientLowGreaterThan50.shape[0],webMSPending_ClientLowGreaterThan100.shape[0],webMSPending_Emergency_ClientCriticalLessThan10.shape[0],webMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],webMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],webMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],webMSPending_Emergency_ClientHighLessThan10.shape[0],webMSPending_Emergency_ClientHighGreaterThan10.shape[0],webMSPending_Emergency_ClientHighGreaterThan50.shape[0],webMSPending_Emergency_ClientHighGreaterThan100.shape[0],webMSPending_Emergency_ClientAverageLessThan10.shape[0],webMSPending_Emergency_ClientAverageGreaterThan10.shape[0],webMSPending_Emergency_ClientAverageGreaterThan50.shape[0],webMSPending_Emergency_ClientAverageGreaterThan100.shape[0],webMSPending_Emergency_ClientLowLessThan10.shape[0],webMSPending_Emergency_ClientLowGreaterThan10.shape[0],webMSPending_Emergency_ClientLowGreaterThan50.shape[0],webMSPending_Emergency_ClientLowGreaterThan100.shape[0],webMSPending_VendorCriticalLessThan10.shape[0],webMSPending_VendorCriticalGreaterThan10.shape[0],webMSPending_VendorCriticalGreaterThan50.shape[0],webMSPending_VendorCriticalGreaterThan100.shape[0],webMSPending_VendorHighLessThan10.shape[0],webMSPending_VendorHighGreaterThan10.shape[0],webMSPending_VendorHighGreaterThan50.shape[0],webMSPending_VendorHighGreaterThan100.shape[0],webMSPending_VendorAverageLessThan10.shape[0],webMSPending_VendorAverageGreaterThan10.shape[0],webMSPending_VendorAverageGreaterThan50.shape[0],webMSPending_VendorAverageGreaterThan100.shape[0],webMSPending_VendorLowLessThan10.shape[0],webMSPending_VendorLowGreaterThan10.shape[0],webMSPending_VendorLowGreaterThan50.shape[0],webMSPending_VendorLowGreaterThan100.shape[0],webMSWork_In_ProgressCriticalLessThan10.shape[0],webMSWork_In_ProgressCriticalGreaterThan10.shape[0],webMSWork_In_ProgressCriticalGreaterThan50.shape[0],webMSWork_In_ProgressCriticalGreaterThan100.shape[0],webMSWork_In_ProgressHighLessThan10.shape[0],webMSWork_In_ProgressHighGreaterThan10.shape[0],webMSWork_In_ProgressHighGreaterThan50.shape[0],webMSWork_In_ProgressHighGreaterThan100.shape[0],webMSWork_In_ProgressAverageLessThan10.shape[0],webMSWork_In_ProgressAverageGreaterThan10.shape[0],webMSWork_In_ProgressAverageGreaterThan50.shape[0],webMSWork_In_ProgressAverageGreaterThan100.shape[0],webMSWork_In_ProgressLowLessThan10.shape[0],webMSWork_In_ProgressLowGreaterThan10.shape[0],webMSWork_In_ProgressLowGreaterThan50.shape[0],webMSWork_In_ProgressLowGreaterThan100.shape[0],webNonMSOpenCriticalLessThan10.shape[0],webNonMSOpenCriticalGreaterThan10.shape[0],webNonMSOpenCriticalGreaterThan50.shape[0],webNonMSOpenCriticalGreaterThan100.shape[0],webNonMSOpenHighLessThan10.shape[0],webNonMSOpenHighGreaterThan10.shape[0],webNonMSOpenHighGreaterThan50.shape[0],webNonMSOpenHighGreaterThan100.shape[0],webNonMSOpenAverageLessThan10.shape[0],webNonMSOpenAverageGreaterThan10.shape[0],webNonMSOpenAverageGreaterThan50.shape[0],webNonMSOpenAverageGreaterThan100.shape[0],webNonMSOpenLowLessThan10.shape[0],webNonMSOpenLowGreaterThan10.shape[0],webNonMSOpenLowGreaterThan50.shape[0],webNonMSOpenLowGreaterThan100.shape[0],webNonMSPending_ClientCriticalLessThan10.shape[0],webNonMSPending_ClientCriticalGreaterThan10.shape[0],webNonMSPending_ClientCriticalGreaterThan50.shape[0],webNonMSPending_ClientCriticalGreaterThan100.shape[0],webNonMSPending_ClientHighLessThan10.shape[0],webNonMSPending_ClientHighGreaterThan10.shape[0],webNonMSPending_ClientHighGreaterThan50.shape[0],webNonMSPending_ClientHighGreaterThan100.shape[0],webNonMSPending_ClientAverageLessThan10.shape[0],webNonMSPending_ClientAverageGreaterThan10.shape[0],webNonMSPending_ClientAverageGreaterThan50.shape[0],webNonMSPending_ClientAverageGreaterThan100.shape[0],webNonMSPending_ClientLowLessThan10.shape[0],webNonMSPending_ClientLowGreaterThan10.shape[0],webNonMSPending_ClientLowGreaterThan50.shape[0],webNonMSPending_ClientLowGreaterThan100.shape[0],webNonMSPending_Emergency_ClientCriticalLessThan10.shape[0],webNonMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],webNonMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],webNonMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],webNonMSPending_Emergency_ClientHighLessThan10.shape[0],webNonMSPending_Emergency_ClientHighGreaterThan10.shape[0],webNonMSPending_Emergency_ClientHighGreaterThan50.shape[0],webNonMSPending_Emergency_ClientHighGreaterThan100.shape[0],webNonMSPending_Emergency_ClientAverageLessThan10.shape[0],webNonMSPending_Emergency_ClientAverageGreaterThan10.shape[0],webNonMSPending_Emergency_ClientAverageGreaterThan50.shape[0],webNonMSPending_Emergency_ClientAverageGreaterThan100.shape[0],webNonMSPending_Emergency_ClientLowLessThan10.shape[0],webNonMSPending_Emergency_ClientLowGreaterThan10.shape[0],webNonMSPending_Emergency_ClientLowGreaterThan50.shape[0],webNonMSPending_Emergency_ClientLowGreaterThan100.shape[0],webNonMSPending_VendorCriticalLessThan10.shape[0],webNonMSPending_VendorCriticalGreaterThan10.shape[0],webNonMSPending_VendorCriticalGreaterThan50.shape[0],webNonMSPending_VendorCriticalGreaterThan100.shape[0],webNonMSPending_VendorHighLessThan10.shape[0],webNonMSPending_VendorHighGreaterThan10.shape[0],webNonMSPending_VendorHighGreaterThan50.shape[0],webNonMSPending_VendorHighGreaterThan100.shape[0],webNonMSPending_VendorAverageLessThan10.shape[0],webNonMSPending_VendorAverageGreaterThan10.shape[0],webNonMSPending_VendorAverageGreaterThan50.shape[0],webNonMSPending_VendorAverageGreaterThan100.shape[0],webNonMSPending_VendorLowLessThan10.shape[0],webNonMSPending_VendorLowGreaterThan10.shape[0],webNonMSPending_VendorLowGreaterThan50.shape[0],webNonMSPending_VendorLowGreaterThan100.shape[0],webNonMSWork_In_ProgressCriticalLessThan10.shape[0],webNonMSWork_In_ProgressCriticalGreaterThan10.shape[0],webNonMSWork_In_ProgressCriticalGreaterThan50.shape[0],webNonMSWork_In_ProgressCriticalGreaterThan100.shape[0],webNonMSWork_In_ProgressHighLessThan10.shape[0],webNonMSWork_In_ProgressHighGreaterThan10.shape[0],webNonMSWork_In_ProgressHighGreaterThan50.shape[0],webNonMSWork_In_ProgressHighGreaterThan100.shape[0],webNonMSWork_In_ProgressAverageLessThan10.shape[0],webNonMSWork_In_ProgressAverageGreaterThan10.shape[0],webNonMSWork_In_ProgressAverageGreaterThan50.shape[0],webNonMSWork_In_ProgressAverageGreaterThan100.shape[0],webNonMSWork_In_ProgressLowLessThan10.shape[0],webNonMSWork_In_ProgressLowGreaterThan10.shape[0],webNonMSWork_In_ProgressLowGreaterThan50.shape[0],webNonMSWork_In_ProgressLowGreaterThan100.shape[0]
]}
rcdf=pd.DataFrame(rowcount)
rcdf.to_excel(rowpath,index=False)
print("success")

#rowcount
#webMSOpenCriticalLessThan10rc=webMSOpenCriticalLessThan10.shape[0]
##(webMSWork_In_ProgressHighLessThan10)
##(Msweb.shape[0])
##(webMSOpen)
##(webMSOpenCritical)













#writer = pd.ExcelWriter(excelPath.replace(".xlsx","1.xlsx"), engine='xlsxwriter')

#web1.to_excel(writer, sheet_name=ng,index=False)
#webpport1.to_excel(writer, sheet_name=es,index=False)
#ediAnalysts1.to_excel(writer, sheet_name=ea,index=False)
#webMethods1.to_excel(writer, sheet_name=we,index=False)

#writer.save()
##("success")
