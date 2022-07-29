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
Ngg=pd.read_excel(excelPath,sheet_name=0,index=False,engine='openpyxl')

#MS_NONMS
MsNgg=Ngg[Ngg['MS_NONMS'] == "MS"]
NonMsNgg=Ngg[Ngg['MS_NONMS'] == "Non_ms"]
#MS Status
NggMSOpen=MsNgg[MsNgg['Status'] == "Open"]
NggMSPending_Client=MsNgg[MsNgg['Status'] == "Pending Client"]
NggMSPending_Emergency_Client=MsNgg[MsNgg['Status'] == "Pending Emergency Client"]
NggMSPending_Vendor=MsNgg[MsNgg['Status'] == "Pending Vendor"]
NggMSWork_In_Progress=MsNgg[MsNgg['Status'] == "Work In Progress"]
#(NggMSWork_In_Progress)
#Non_MS Status
NggNonMSOpen=NonMsNgg[NonMsNgg['Status'] == "Open"]
NggNonMSPending_Client=NonMsNgg[NonMsNgg['Status'] == "Pending Client"]
NggNonMSPending_Emergency_Client=NonMsNgg[NonMsNgg['Status'] == "Pending Emergency Client"]
NggNonMSPending_Vendor=NonMsNgg[NonMsNgg['Status'] == "Pending Vendor"]
NggNonMSWork_In_Progress=NonMsNgg[NonMsNgg['Status'] == "Work In Progress"]

#MS_OPEN_PRIORITY**
NggMSOpenCritical=NggMSOpen[NggMSOpen['Priority'] == 1]
NggMSOpenHigh=NggMSOpen[NggMSOpen['Priority'] == 2]
NggMSOpenAverage=NggMSOpen[NggMSOpen['Priority'] == 3]
NggMSOpenLow=NggMSOpen[NggMSOpen['Priority'] == 4]

##MS_Pending_Client_PRIORITY
NggMSPending_ClientCritical=NggMSPending_Client[NggMSPending_Client['Priority'] == 1]
NggMSPending_ClientHigh=NggMSPending_Client[NggMSPending_Client['Priority'] == 2]
NggMSPending_ClientAverage=NggMSPending_Client[NggMSPending_Client['Priority'] == 3]
NggMSPending_ClientLow=NggMSPending_Client[NggMSPending_Client['Priority'] == 4]

#MS_Pending_Emergency_Client
NggMSPending_Emergency_ClientCritical=NggMSPending_Emergency_Client[NggMSPending_Emergency_Client['Priority'] == 1]
NggMSPending_Emergency_ClientHigh=NggMSPending_Emergency_Client[NggMSPending_Emergency_Client['Priority'] == 2]
NggMSPending_Emergency_ClientAverage=NggMSPending_Emergency_Client[NggMSPending_Emergency_Client['Priority'] == 3]
NggMSPending_Emergency_ClientLow=NggMSPending_Emergency_Client[NggMSPending_Emergency_Client['Priority'] == 4]

#MS_Pending_Vendor
NggMSPending_VendorCritical=NggMSPending_Vendor[NggMSPending_Vendor['Priority'] == 1]
NggMSPending_VendorHigh=NggMSPending_Vendor[NggMSPending_Vendor['Priority'] == 2]
NggMSPending_VendorAverage=NggMSPending_Vendor[NggMSPending_Vendor['Priority'] == 3]
NggMSPending_VendorLow=NggMSPending_Vendor[NggMSPending_Vendor['Priority'] == 4]

#MS_Work_In_Progress
NggMSWork_In_ProgressCritical=NggMSWork_In_Progress[NggMSWork_In_Progress['Priority'] == 1]
NggMSWork_In_ProgressHigh=NggMSWork_In_Progress[NggMSWork_In_Progress['Priority'] == 2]
NggMSWork_In_ProgressAverage=NggMSWork_In_Progress[NggMSWork_In_Progress['Priority'] == 3]
NggMSWork_In_ProgressLow=NggMSWork_In_Progress[NggMSWork_In_Progress['Priority'] == 4]
#(NggMSWork_In_ProgressHigh)


#NonMS_OPEN_PRIORITY
NggNonMSOpenCritical=NggNonMSOpen[NggNonMSOpen['Priority'] == 1]
NggNonMSOpenHigh=NggNonMSOpen[NggNonMSOpen['Priority'] == 2]
NggNonMSOpenAverage=NggNonMSOpen[NggNonMSOpen['Priority'] == 3]
NggNonMSOpenLow=NggNonMSOpen[NggNonMSOpen['Priority'] == 4]

##NonMS_Pending_Client_PRIORITY
NggNonMSPending_ClientCritical=NggNonMSPending_Client[NggNonMSPending_Client['Priority'] == 1]
NggNonMSPending_ClientHigh=NggNonMSPending_Client[NggNonMSPending_Client['Priority'] == 2]
NggNonMSPending_ClientAverage=NggNonMSPending_Client[NggNonMSPending_Client['Priority'] == 3]
NggNonMSPending_ClientLow=NggNonMSPending_Client[NggNonMSPending_Client['Priority'] == 4]

#NonMS_Pending_Emergency_Client
NggNonMSPending_Emergency_ClientCritical=NggNonMSPending_Emergency_Client[NggNonMSPending_Emergency_Client['Priority'] == 1]
NggNonMSPending_Emergency_ClientHigh=NggNonMSPending_Emergency_Client[NggNonMSPending_Emergency_Client['Priority'] == 2]
NggNonMSPending_Emergency_ClientAverage=NggNonMSPending_Emergency_Client[NggNonMSPending_Emergency_Client['Priority'] == 3]
NggNonMSPending_Emergency_ClientLow=NggNonMSPending_Emergency_Client[NggNonMSPending_Emergency_Client['Priority'] == 4]

#NonMS_Pending_Vendor
NggNonMSPending_VendorCritical=NggNonMSPending_Vendor[NggNonMSPending_Vendor['Priority'] == 1]
NggNonMSPending_VendorHigh=NggNonMSPending_Vendor[NggNonMSPending_Vendor['Priority'] == 2]
NggNonMSPending_VendorAverage=NggNonMSPending_Vendor[NggNonMSPending_Vendor['Priority'] == 3]
NggNonMSPending_VendorLow=NggNonMSPending_Vendor[NggNonMSPending_Vendor['Priority'] == 4]

#NonMS_Work_In_Progress
NggNonMSWork_In_ProgressCritical=NggNonMSWork_In_Progress[NggNonMSWork_In_Progress['Priority'] == 1]
NggNonMSWork_In_ProgressHigh=NggNonMSWork_In_Progress[NggNonMSWork_In_Progress['Priority'] == 2]
NggNonMSWork_In_ProgressAverage=NggNonMSWork_In_Progress[NggNonMSWork_In_Progress['Priority'] == 3]
NggNonMSWork_In_ProgressLow=NggNonMSWork_In_Progress[NggNonMSWork_In_Progress['Priority'] == 4]

#NggMSOpenCriticalAging
NggMSOpenCriticalLessThan10=NggMSOpenCritical[NggMSOpenCritical['Ageing days'] <=10]
NggMSOpenCriticalGreaterThan10=NggMSOpenCritical[(NggMSOpenCritical['Ageing days'] > 10)&(NggMSOpenCritical['Ageing days'] <=50)]
NggMSOpenCriticalGreaterThan50=NggMSOpenCritical[(NggMSOpenCritical['Ageing days'] > 50)&(NggMSOpenCritical['Ageing days'] <=100)]
NggMSOpenCriticalGreaterThan100=NggMSOpenCritical[NggMSOpenCritical['Ageing days'] > 100]

#NggMSOpenHighAging
NggMSOpenHighLessThan10=NggMSOpenHigh[NggMSOpenHigh['Ageing days'] <=10]
NggMSOpenHighGreaterThan10=NggMSOpenHigh[(NggMSOpenHigh['Ageing days'] > 10)&(NggMSOpenHigh['Ageing days'] <=50)]
NggMSOpenHighGreaterThan50=NggMSOpenHigh[(NggMSOpenHigh['Ageing days'] > 50)&(NggMSOpenHigh['Ageing days'] <=100)]
NggMSOpenHighGreaterThan100=NggMSOpenHigh[NggMSOpenHigh['Ageing days'] > 100]

#NggMSOpenAverageAging
NggMSOpenAverageLessThan10=NggMSOpenAverage[NggMSOpenAverage['Ageing days'] <=10]
NggMSOpenAverageGreaterThan10=NggMSOpenAverage[(NggMSOpenAverage['Ageing days'] > 10)&(NggMSOpenAverage['Ageing days'] <=50)]
NggMSOpenAverageGreaterThan50=NggMSOpenAverage[(NggMSOpenAverage['Ageing days'] > 50)&(NggMSOpenAverage['Ageing days'] <=100)]
NggMSOpenAverageGreaterThan100=NggMSOpenAverage[NggMSOpenAverage['Ageing days'] > 100]

#NggMSOpenLowAging
NggMSOpenLowLessThan10=NggMSOpenLow[NggMSOpenLow['Ageing days'] <=10]
NggMSOpenLowGreaterThan10=NggMSOpenLow[(NggMSOpenLow['Ageing days'] > 10)&(NggMSOpenLow['Ageing days'] <=50)]
NggMSOpenLowGreaterThan50=NggMSOpenLow[(NggMSOpenLow['Ageing days'] > 50)&(NggMSOpenLow['Ageing days'] <=100)]
NggMSOpenLowGreaterThan100=NggMSOpenLow[NggMSOpenLow['Ageing days'] > 100]


#NggMSPending_ClientCriticalAging
NggMSPending_ClientCriticalLessThan10=NggMSPending_ClientCritical[NggMSPending_ClientCritical['Ageing days'] <=10]
NggMSPending_ClientCriticalGreaterThan10=NggMSPending_ClientCritical[(NggMSPending_ClientCritical['Ageing days'] > 10)&(NggMSPending_ClientCritical['Ageing days'] <=50)]
NggMSPending_ClientCriticalGreaterThan50=NggMSPending_ClientCritical[(NggMSPending_ClientCritical['Ageing days'] > 50)&(NggMSPending_ClientCritical['Ageing days'] <=100)]
NggMSPending_ClientCriticalGreaterThan100=NggMSPending_ClientCritical[NggMSPending_ClientCritical['Ageing days'] > 100]

#NggMSPending_ClientHighAging
NggMSPending_ClientHighLessThan10=NggMSPending_ClientHigh[NggMSPending_ClientHigh['Ageing days'] <=10]
NggMSPending_ClientHighGreaterThan10=NggMSPending_ClientHigh[(NggMSPending_ClientHigh['Ageing days'] > 10)&(NggMSPending_ClientHigh['Ageing days'] <=50)]
NggMSPending_ClientHighGreaterThan50=NggMSPending_ClientHigh[(NggMSPending_ClientHigh['Ageing days'] > 50)&(NggMSPending_ClientHigh['Ageing days'] <=100)]
NggMSPending_ClientHighGreaterThan100=NggMSPending_ClientHigh[NggMSPending_ClientHigh['Ageing days'] > 100]

#NggMSPending_ClientAverageAging
NggMSPending_ClientAverageLessThan10=NggMSPending_ClientAverage[NggMSPending_ClientAverage['Ageing days'] <=10]
NggMSPending_ClientAverageGreaterThan10=NggMSPending_ClientAverage[(NggMSPending_ClientAverage['Ageing days'] > 10)&(NggMSPending_ClientAverage['Ageing days'] <=50)]
NggMSPending_ClientAverageGreaterThan50=NggMSPending_ClientAverage[(NggMSPending_ClientAverage['Ageing days'] > 50)&(NggMSPending_ClientAverage['Ageing days'] <=100)]
NggMSPending_ClientAverageGreaterThan100=NggMSPending_ClientAverage[NggMSPending_ClientAverage['Ageing days'] > 100]

#NggMSPending_ClientLowAging
NggMSPending_ClientLowLessThan10=NggMSPending_ClientLow[NggMSPending_ClientLow['Ageing days'] <=10]
NggMSPending_ClientLowGreaterThan10=NggMSPending_ClientLow[(NggMSPending_ClientLow['Ageing days'] > 10)&(NggMSPending_ClientLow['Ageing days'] <=50)]
NggMSPending_ClientLowGreaterThan50=NggMSPending_ClientLow[(NggMSPending_ClientLow['Ageing days'] > 50)&(NggMSPending_ClientLow['Ageing days'] <=100)]
NggMSPending_ClientLowGreaterThan100=NggMSPending_ClientLow[NggMSPending_ClientLow['Ageing days'] > 100]


#NggMSPending_Emergency_ClientCriticalAging
NggMSPending_Emergency_ClientCriticalLessThan10=NggMSPending_Emergency_ClientCritical[NggMSPending_Emergency_ClientCritical['Ageing days'] <=10]
NggMSPending_Emergency_ClientCriticalGreaterThan10=NggMSPending_Emergency_ClientCritical[(NggMSPending_Emergency_ClientCritical['Ageing days'] > 10)&(NggMSPending_Emergency_ClientCritical['Ageing days'] <=50)]
NggMSPending_Emergency_ClientCriticalGreaterThan50=NggMSPending_Emergency_ClientCritical[(NggMSPending_Emergency_ClientCritical['Ageing days'] > 50)&(NggMSPending_Emergency_ClientCritical['Ageing days'] <=100)]
NggMSPending_Emergency_ClientCriticalGreaterThan100=NggMSPending_Emergency_ClientCritical[NggMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#NggMSPending_Emergency_ClientHighAging
NggMSPending_Emergency_ClientHighLessThan10=NggMSPending_Emergency_ClientHigh[NggMSPending_Emergency_ClientHigh['Ageing days'] <=10]
NggMSPending_Emergency_ClientHighGreaterThan10=NggMSPending_Emergency_ClientHigh[(NggMSPending_Emergency_ClientHigh['Ageing days'] > 10)&(NggMSPending_Emergency_ClientHigh['Ageing days'] <=50)]
NggMSPending_Emergency_ClientHighGreaterThan50=NggMSPending_Emergency_ClientHigh[(NggMSPending_Emergency_ClientHigh['Ageing days'] > 50)&(NggMSPending_Emergency_ClientHigh['Ageing days'] <=100)]
NggMSPending_Emergency_ClientHighGreaterThan100=NggMSPending_Emergency_ClientHigh[NggMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#NggMSPending_Emergency_ClientAverageAging
NggMSPending_Emergency_ClientAverageLessThan10=NggMSPending_Emergency_ClientAverage[NggMSPending_Emergency_ClientAverage['Ageing days'] <=10]
NggMSPending_Emergency_ClientAverageGreaterThan10=NggMSPending_Emergency_ClientAverage[(NggMSPending_Emergency_ClientAverage['Ageing days'] > 10)&(NggMSPending_Emergency_ClientAverage['Ageing days'] <=50)]
NggMSPending_Emergency_ClientAverageGreaterThan50=NggMSPending_Emergency_ClientAverage[(NggMSPending_Emergency_ClientAverage['Ageing days'] > 50)&(NggMSPending_Emergency_ClientAverage['Ageing days'] <=100)]
NggMSPending_Emergency_ClientAverageGreaterThan100=NggMSPending_Emergency_ClientAverage[NggMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#NggMSPending_Emergency_ClientLowAging
NggMSPending_Emergency_ClientLowLessThan10=NggMSPending_Emergency_ClientLow[NggMSPending_Emergency_ClientLow['Ageing days'] <=10]
NggMSPending_Emergency_ClientLowGreaterThan10=NggMSPending_Emergency_ClientLow[(NggMSPending_Emergency_ClientLow['Ageing days'] > 10)&(NggMSPending_Emergency_ClientLow['Ageing days'] <=50)]
NggMSPending_Emergency_ClientLowGreaterThan50=NggMSPending_Emergency_ClientLow[(NggMSPending_Emergency_ClientLow['Ageing days'] > 50)&(NggMSPending_Emergency_ClientLow['Ageing days'] <=100)]
NggMSPending_Emergency_ClientLowGreaterThan100=NggMSPending_Emergency_ClientLow[NggMSPending_Emergency_ClientLow['Ageing days'] > 100]


#NggMSPending_VendorCriticalAging
NggMSPending_VendorCriticalLessThan10=NggMSPending_VendorCritical[NggMSPending_VendorCritical['Ageing days'] <=10]
NggMSPending_VendorCriticalGreaterThan10=NggMSPending_VendorCritical[(NggMSPending_VendorCritical['Ageing days'] > 10)&(NggMSPending_VendorCritical['Ageing days'] <=50)]
NggMSPending_VendorCriticalGreaterThan50=NggMSPending_VendorCritical[(NggMSPending_VendorCritical['Ageing days'] > 50)&(NggMSPending_VendorCritical['Ageing days'] <=100)]
NggMSPending_VendorCriticalGreaterThan100=NggMSPending_VendorCritical[NggMSPending_VendorCritical['Ageing days'] > 100]

#NggMSPending_VendorHighAging
NggMSPending_VendorHighLessThan10=NggMSPending_VendorHigh[NggMSPending_VendorHigh['Ageing days'] <=10]
NggMSPending_VendorHighGreaterThan10=NggMSPending_VendorHigh[(NggMSPending_VendorHigh['Ageing days'] > 10)&(NggMSPending_VendorHigh['Ageing days'] <=50)]
NggMSPending_VendorHighGreaterThan50=NggMSPending_VendorHigh[(NggMSPending_VendorHigh['Ageing days'] > 50)&(NggMSPending_VendorHigh['Ageing days'] <=100)]
NggMSPending_VendorHighGreaterThan100=NggMSPending_VendorHigh[NggMSPending_VendorHigh['Ageing days'] > 100]

#NggMSPending_VendorAverageAging
NggMSPending_VendorAverageLessThan10=NggMSPending_VendorAverage[NggMSPending_VendorAverage['Ageing days'] <=10]
NggMSPending_VendorAverageGreaterThan10=NggMSPending_VendorAverage[(NggMSPending_VendorAverage['Ageing days'] > 10)&(NggMSPending_VendorAverage['Ageing days'] <=50)]
NggMSPending_VendorAverageGreaterThan50=NggMSPending_VendorAverage[(NggMSPending_VendorAverage['Ageing days'] > 50)&(NggMSPending_VendorAverage['Ageing days'] <=100)]
NggMSPending_VendorAverageGreaterThan100=NggMSPending_VendorAverage[NggMSPending_VendorAverage['Ageing days'] > 100]

#NggMSPending_VendorLowAging
NggMSPending_VendorLowLessThan10=NggMSPending_VendorLow[NggMSPending_VendorLow['Ageing days'] <=10]
NggMSPending_VendorLowGreaterThan10=NggMSPending_VendorLow[(NggMSPending_VendorLow['Ageing days'] > 10)&(NggMSPending_VendorLow['Ageing days'] <=50)]
NggMSPending_VendorLowGreaterThan50=NggMSPending_VendorLow[(NggMSPending_VendorLow['Ageing days'] > 50)&(NggMSPending_VendorLow['Ageing days'] <=100)]
NggMSPending_VendorLowGreaterThan100=NggMSPending_VendorLow[NggMSPending_VendorLow['Ageing days'] > 100]


#NggMSWork_In_ProgressCriticalAging
NggMSWork_In_ProgressCriticalLessThan10=NggMSWork_In_ProgressCritical[NggMSWork_In_ProgressCritical['Ageing days'] <=10]
NggMSWork_In_ProgressCriticalGreaterThan10=NggMSWork_In_ProgressCritical[(NggMSWork_In_ProgressCritical['Ageing days'] > 10)&(NggMSWork_In_ProgressCritical['Ageing days'] <=50)]
NggMSWork_In_ProgressCriticalGreaterThan50=NggMSWork_In_ProgressCritical[(NggMSWork_In_ProgressCritical['Ageing days'] > 50)&(NggMSWork_In_ProgressCritical['Ageing days'] <=100)]
NggMSWork_In_ProgressCriticalGreaterThan100=NggMSWork_In_ProgressCritical[NggMSWork_In_ProgressCritical['Ageing days'] > 100]

#NggMSWork_In_ProgressHighAging
NggMSWork_In_ProgressHighLessThan10=NggMSWork_In_ProgressHigh[NggMSWork_In_ProgressHigh['Ageing days'] <=10]
NggMSWork_In_ProgressHighGreaterThan10=NggMSWork_In_ProgressHigh[(NggMSWork_In_ProgressHigh['Ageing days'] > 10)&(NggMSWork_In_ProgressHigh['Ageing days'] <=50)]
NggMSWork_In_ProgressHighGreaterThan50=NggMSWork_In_ProgressHigh[(NggMSWork_In_ProgressHigh['Ageing days'] > 50)&(NggMSWork_In_ProgressHigh['Ageing days'] <=100)]
NggMSWork_In_ProgressHighGreaterThan100=NggMSWork_In_ProgressHigh[NggMSWork_In_ProgressHigh['Ageing days'] > 100]
#(NggMSWork_In_ProgressHighLessThan10)
#(NggMSWork_In_ProgressHighLessThan10.shape[0])
#NggMSWork_In_ProgressAverageAging
NggMSWork_In_ProgressAverageLessThan10=NggMSWork_In_ProgressAverage[NggMSWork_In_ProgressAverage['Ageing days'] <=10]
NggMSWork_In_ProgressAverageGreaterThan10=NggMSWork_In_ProgressAverage[(NggMSWork_In_ProgressAverage['Ageing days'] > 10)&(NggMSWork_In_ProgressAverage['Ageing days'] <=50)]
NggMSWork_In_ProgressAverageGreaterThan50=NggMSWork_In_ProgressAverage[(NggMSWork_In_ProgressAverage['Ageing days'] > 50)&(NggMSWork_In_ProgressAverage['Ageing days'] <=100)]
NggMSWork_In_ProgressAverageGreaterThan100=NggMSWork_In_ProgressAverage[NggMSWork_In_ProgressAverage['Ageing days'] > 100]

#NggMSWork_In_ProgressLowAging
NggMSWork_In_ProgressLowLessThan10=NggMSWork_In_ProgressLow[NggMSWork_In_ProgressLow['Ageing days'] <=10]
NggMSWork_In_ProgressLowGreaterThan10=NggMSWork_In_ProgressLow[(NggMSWork_In_ProgressLow['Ageing days'] > 10)&(NggMSWork_In_ProgressLow['Ageing days'] <=50)]
NggMSWork_In_ProgressLowGreaterThan50=NggMSWork_In_ProgressLow[(NggMSWork_In_ProgressLow['Ageing days'] > 50)&(NggMSWork_In_ProgressLow['Ageing days'] <=100)]
NggMSWork_In_ProgressLowGreaterThan100=NggMSWork_In_ProgressLow[NggMSWork_In_ProgressLow['Ageing days'] > 100]



#NggNonMSOpenCriticalAging
NggNonMSOpenCriticalLessThan10=NggNonMSOpenCritical[NggNonMSOpenCritical['Ageing days'] <=10]
NggNonMSOpenCriticalGreaterThan10=NggNonMSOpenCritical[(NggNonMSOpenCritical['Ageing days'] > 10)&(NggNonMSOpenCritical['Ageing days'] <=50)]
NggNonMSOpenCriticalGreaterThan50=NggNonMSOpenCritical[(NggNonMSOpenCritical['Ageing days'] > 50)&(NggNonMSOpenCritical['Ageing days'] <=100)]
NggNonMSOpenCriticalGreaterThan100=NggNonMSOpenCritical[NggNonMSOpenCritical['Ageing days'] > 100]

#NggNonMSOpenHighAging
NggNonMSOpenHighLessThan10=NggNonMSOpenHigh[NggNonMSOpenHigh['Ageing days'] <=10]
NggNonMSOpenHighGreaterThan10=NggNonMSOpenHigh[(NggNonMSOpenHigh['Ageing days'] > 10)&(NggNonMSOpenHigh['Ageing days'] <=50)]
NggNonMSOpenHighGreaterThan50=NggNonMSOpenHigh[(NggNonMSOpenHigh['Ageing days'] > 50)&(NggNonMSOpenHigh['Ageing days'] <=100)]
NggNonMSOpenHighGreaterThan100=NggNonMSOpenHigh[NggNonMSOpenHigh['Ageing days'] > 100]

#NggNonMSOpenAverageAging
NggNonMSOpenAverageLessThan10=NggNonMSOpenAverage[NggNonMSOpenAverage['Ageing days'] <=10]
NggNonMSOpenAverageGreaterThan10=NggNonMSOpenAverage[(NggNonMSOpenAverage['Ageing days'] > 10)&(NggNonMSOpenAverage['Ageing days'] <=50)]
NggNonMSOpenAverageGreaterThan50=NggNonMSOpenAverage[(NggNonMSOpenAverage['Ageing days'] > 50)&(NggNonMSOpenAverage['Ageing days'] <= 100)]
NggNonMSOpenAverageGreaterThan100=NggNonMSOpenAverage[NggNonMSOpenAverage['Ageing days'] > 100]

#NggNonMSOpenLowAging
NggNonMSOpenLowLessThan10=NggNonMSOpenLow[NggNonMSOpenLow['Ageing days'] <=10]
NggNonMSOpenLowGreaterThan10=NggNonMSOpenLow[(NggNonMSOpenLow['Ageing days'] > 10)&(NggNonMSOpenLow['Ageing days'] <=50)]
NggNonMSOpenLowGreaterThan50=NggNonMSOpenLow[(NggNonMSOpenLow['Ageing days'] > 50)&(NggNonMSOpenLow['Ageing days'] <=100)]
NggNonMSOpenLowGreaterThan100=NggNonMSOpenLow[NggNonMSOpenLow['Ageing days'] > 100]


#NggNonMSPending_ClientCriticalAging
NggNonMSPending_ClientCriticalLessThan10=NggNonMSPending_ClientCritical[NggNonMSPending_ClientCritical['Ageing days'] <=10]
NggNonMSPending_ClientCriticalGreaterThan10=NggNonMSPending_ClientCritical[(NggNonMSPending_ClientCritical['Ageing days'] > 10)&(NggNonMSPending_ClientCritical['Ageing days'] <=50)]
NggNonMSPending_ClientCriticalGreaterThan50=NggNonMSPending_ClientCritical[(NggNonMSPending_ClientCritical['Ageing days'] > 50)&(NggNonMSPending_ClientCritical['Ageing days'] <=100)]
NggNonMSPending_ClientCriticalGreaterThan100=NggNonMSPending_ClientCritical[NggNonMSPending_ClientCritical['Ageing days'] > 100]

#NggNonMSPending_ClientHighAging
NggNonMSPending_ClientHighLessThan10=NggNonMSPending_ClientHigh[NggNonMSPending_ClientHigh['Ageing days'] <=10]
NggNonMSPending_ClientHighGreaterThan10=NggNonMSPending_ClientHigh[(NggNonMSPending_ClientHigh['Ageing days'] > 10) & (NggNonMSPending_ClientHigh['Ageing days'] <= 50)]
NggNonMSPending_ClientHighGreaterThan50=NggNonMSPending_ClientHigh[(NggNonMSPending_ClientHigh['Ageing days'] > 50) & (NggNonMSPending_ClientHigh['Ageing days'] <=100)]
NggNonMSPending_ClientHighGreaterThan100=NggNonMSPending_ClientHigh[NggNonMSPending_ClientHigh['Ageing days'] > 100]

#NggNonMSPending_ClientAverageAging
NggNonMSPending_ClientAverageLessThan10=NggNonMSPending_ClientAverage[NggNonMSPending_ClientAverage['Ageing days'] <=10]
NggNonMSPending_ClientAverageGreaterThan10=NggNonMSPending_ClientAverage[(NggNonMSPending_ClientAverage['Ageing days'] > 10) & (NggNonMSPending_ClientAverage['Ageing days'] <=50)]
NggNonMSPending_ClientAverageGreaterThan50=NggNonMSPending_ClientAverage[(NggNonMSPending_ClientAverage['Ageing days'] > 50) & (NggNonMSPending_ClientAverage['Ageing days'] <=100)]
NggNonMSPending_ClientAverageGreaterThan100=NggNonMSPending_ClientAverage[NggNonMSPending_ClientAverage['Ageing days'] > 100]

#NggNonMSPending_ClientLowAging
NggNonMSPending_ClientLowLessThan10=NggNonMSPending_ClientLow[NggNonMSPending_ClientLow['Ageing days'] <=10]
NggNonMSPending_ClientLowGreaterThan10=NggNonMSPending_ClientLow[(NggNonMSPending_ClientLow['Ageing days'] > 10)&(NggNonMSPending_ClientLow['Ageing days'] <=50)]
NggNonMSPending_ClientLowGreaterThan50=NggNonMSPending_ClientLow[(NggNonMSPending_ClientLow['Ageing days'] > 50)& (NggNonMSPending_ClientLow['Ageing days'] <=100)]
NggNonMSPending_ClientLowGreaterThan100=NggNonMSPending_ClientLow[NggNonMSPending_ClientLow['Ageing days'] > 100]


#NggNonMSPending_Emergency_ClientCriticalAging
NggNonMSPending_Emergency_ClientCriticalLessThan10=NggNonMSPending_Emergency_ClientCritical[NggNonMSPending_Emergency_ClientCritical['Ageing days'] <=10]
NggNonMSPending_Emergency_ClientCriticalGreaterThan10=NggNonMSPending_Emergency_ClientCritical[(NggNonMSPending_Emergency_ClientCritical['Ageing days'] > 10)&(NggNonMSPending_Emergency_ClientCritical['Ageing days'] <=50)]
NggNonMSPending_Emergency_ClientCriticalGreaterThan50=NggNonMSPending_Emergency_ClientCritical[(NggNonMSPending_Emergency_ClientCritical['Ageing days'] > 50)&(NggNonMSPending_Emergency_ClientCritical['Ageing days'] <=100)]
NggNonMSPending_Emergency_ClientCriticalGreaterThan100=NggNonMSPending_Emergency_ClientCritical[NggNonMSPending_Emergency_ClientCritical['Ageing days'] > 100]

#NggNonMSPending_Emergency_ClientHighAging
NggNonMSPending_Emergency_ClientHighLessThan10=NggNonMSPending_Emergency_ClientHigh[NggNonMSPending_Emergency_ClientHigh['Ageing days'] <=10]
NggNonMSPending_Emergency_ClientHighGreaterThan10=NggNonMSPending_Emergency_ClientHigh[(NggNonMSPending_Emergency_ClientHigh['Ageing days'] > 10)&(NggNonMSPending_Emergency_ClientHigh['Ageing days'] <=50)]
NggNonMSPending_Emergency_ClientHighGreaterThan50=NggNonMSPending_Emergency_ClientHigh[(NggNonMSPending_Emergency_ClientHigh['Ageing days'] > 50)&(NggNonMSPending_Emergency_ClientHigh['Ageing days'] <=100)]
NggNonMSPending_Emergency_ClientHighGreaterThan100=NggNonMSPending_Emergency_ClientHigh[NggNonMSPending_Emergency_ClientHigh['Ageing days'] > 100]

#NggNonMSPending_Emergency_ClientAverageAging
NggNonMSPending_Emergency_ClientAverageLessThan10=NggNonMSPending_Emergency_ClientAverage[NggNonMSPending_Emergency_ClientAverage['Ageing days'] <=10]
NggNonMSPending_Emergency_ClientAverageGreaterThan10=NggNonMSPending_Emergency_ClientAverage[(NggNonMSPending_Emergency_ClientAverage['Ageing days'] > 10)&(NggNonMSPending_Emergency_ClientAverage['Ageing days'] <=50)]
NggNonMSPending_Emergency_ClientAverageGreaterThan50=NggNonMSPending_Emergency_ClientAverage[(NggNonMSPending_Emergency_ClientAverage['Ageing days'] > 50)&(NggNonMSPending_Emergency_ClientAverage['Ageing days'] <=100)]
NggNonMSPending_Emergency_ClientAverageGreaterThan100=NggNonMSPending_Emergency_ClientAverage[NggNonMSPending_Emergency_ClientAverage['Ageing days'] > 100]

#NggNonMSPending_Emergency_ClientLowAging
NggNonMSPending_Emergency_ClientLowLessThan10=NggNonMSPending_Emergency_ClientLow[NggNonMSPending_Emergency_ClientLow['Ageing days'] <=10]
NggNonMSPending_Emergency_ClientLowGreaterThan10=NggNonMSPending_Emergency_ClientLow[(NggNonMSPending_Emergency_ClientLow['Ageing days'] > 10)&(NggNonMSPending_Emergency_ClientLow['Ageing days'] <=50)]
NggNonMSPending_Emergency_ClientLowGreaterThan50=NggNonMSPending_Emergency_ClientLow[(NggNonMSPending_Emergency_ClientLow['Ageing days'] > 50)&(NggNonMSPending_Emergency_ClientLow['Ageing days'] <=100)]
NggNonMSPending_Emergency_ClientLowGreaterThan100=NggNonMSPending_Emergency_ClientLow[NggNonMSPending_Emergency_ClientLow['Ageing days'] > 100]


#NggNonMSPending_VendorCriticalAging
NggNonMSPending_VendorCriticalLessThan10=NggNonMSPending_VendorCritical[NggNonMSPending_VendorCritical['Ageing days'] <=10]
NggNonMSPending_VendorCriticalGreaterThan10=NggNonMSPending_VendorCritical[(NggNonMSPending_VendorCritical['Ageing days'] > 10)&(NggNonMSPending_VendorCritical['Ageing days'] <=50)]
NggNonMSPending_VendorCriticalGreaterThan50=NggNonMSPending_VendorCritical[(NggNonMSPending_VendorCritical['Ageing days'] > 50)&(NggNonMSPending_VendorCritical['Ageing days'] <=100)]
NggNonMSPending_VendorCriticalGreaterThan100=NggNonMSPending_VendorCritical[NggNonMSPending_VendorCritical['Ageing days'] > 100]

#NggNonMSPending_VendorHighAging
NggNonMSPending_VendorHighLessThan10=NggNonMSPending_VendorHigh[NggNonMSPending_VendorHigh['Ageing days'] <=10]
NggNonMSPending_VendorHighGreaterThan10=NggNonMSPending_VendorHigh[(NggNonMSPending_VendorHigh['Ageing days'] > 10)&(NggNonMSPending_VendorHigh['Ageing days'] <=50)]
NggNonMSPending_VendorHighGreaterThan50=NggNonMSPending_VendorHigh[(NggNonMSPending_VendorHigh['Ageing days'] > 50)&(NggNonMSPending_VendorHigh['Ageing days'] <=100)]
NggNonMSPending_VendorHighGreaterThan100=NggNonMSPending_VendorHigh[NggNonMSPending_VendorHigh['Ageing days'] > 100]

#NggNonMSPending_VendorAverageAging
NggNonMSPending_VendorAverageLessThan10=NggNonMSPending_VendorAverage[NggNonMSPending_VendorAverage['Ageing days'] <=10]
NggNonMSPending_VendorAverageGreaterThan10=NggNonMSPending_VendorAverage[(NggNonMSPending_VendorAverage['Ageing days'] > 10)&(NggNonMSPending_VendorAverage['Ageing days'] <=50)]
NggNonMSPending_VendorAverageGreaterThan50=NggNonMSPending_VendorAverage[(NggNonMSPending_VendorAverage['Ageing days'] > 50)&(NggNonMSPending_VendorAverage['Ageing days'] <=100)]
NggNonMSPending_VendorAverageGreaterThan100=NggNonMSPending_VendorAverage[NggNonMSPending_VendorAverage['Ageing days'] > 100]

#NggNonMSPending_VendorLowAging
NggNonMSPending_VendorLowLessThan10=NggNonMSPending_VendorLow[NggNonMSPending_VendorLow['Ageing days'] <=10]
NggNonMSPending_VendorLowGreaterThan10=NggNonMSPending_VendorLow[(NggNonMSPending_VendorLow['Ageing days'] > 10)&(NggNonMSPending_VendorLow['Ageing days'] <=50)]
NggNonMSPending_VendorLowGreaterThan50=NggNonMSPending_VendorLow[(NggNonMSPending_VendorLow['Ageing days'] > 50)&(NggNonMSPending_VendorLow['Ageing days'] <=100)]
NggNonMSPending_VendorLowGreaterThan100=NggNonMSPending_VendorLow[NggNonMSPending_VendorLow['Ageing days'] > 100]


#NggNonMSWork_In_ProgressCriticalAging
NggNonMSWork_In_ProgressCriticalLessThan10=NggNonMSWork_In_ProgressCritical[NggNonMSWork_In_ProgressCritical['Ageing days'] <=10]
NggNonMSWork_In_ProgressCriticalGreaterThan10=NggNonMSWork_In_ProgressCritical[(NggNonMSWork_In_ProgressCritical['Ageing days'] > 10)&(NggNonMSWork_In_ProgressCritical['Ageing days']<=50)]
NggNonMSWork_In_ProgressCriticalGreaterThan50=NggNonMSWork_In_ProgressCritical[(NggNonMSWork_In_ProgressCritical['Ageing days'] > 50)&(NggNonMSWork_In_ProgressCritical['Ageing days'] <=100)]
NggNonMSWork_In_ProgressCriticalGreaterThan100=NggNonMSWork_In_ProgressCritical[NggNonMSWork_In_ProgressCritical['Ageing days'] > 100]

#NggNonMSWork_In_ProgressHighAging
NggNonMSWork_In_ProgressHighLessThan10=NggNonMSWork_In_ProgressHigh[NggNonMSWork_In_ProgressHigh['Ageing days'] <=10]
NggNonMSWork_In_ProgressHighGreaterThan10=NggNonMSWork_In_ProgressHigh[(NggNonMSWork_In_ProgressHigh['Ageing days'] > 10)&(NggNonMSWork_In_ProgressHigh['Ageing days'] <=50)]
NggNonMSWork_In_ProgressHighGreaterThan50=NggNonMSWork_In_ProgressHigh[(NggNonMSWork_In_ProgressHigh['Ageing days'] > 50)&(NggNonMSWork_In_ProgressHigh['Ageing days'] <=100)]
NggNonMSWork_In_ProgressHighGreaterThan100=NggNonMSWork_In_ProgressHigh[NggNonMSWork_In_ProgressHigh['Ageing days'] > 100]

#NggNonMSWork_In_ProgressAverageAging
NggNonMSWork_In_ProgressAverageLessThan10=NggNonMSWork_In_ProgressAverage[NggNonMSWork_In_ProgressAverage['Ageing days'] <=10]
NggNonMSWork_In_ProgressAverageGreaterThan10=NggNonMSWork_In_ProgressAverage[(NggNonMSWork_In_ProgressAverage['Ageing days'] > 10) &(NggNonMSWork_In_ProgressAverage['Ageing days'] <=50)]
NggNonMSWork_In_ProgressAverageGreaterThan50=NggNonMSWork_In_ProgressAverage[(NggNonMSWork_In_ProgressAverage['Ageing days'] > 50) &(NggNonMSWork_In_ProgressAverage['Ageing days'] <=100)]
NggNonMSWork_In_ProgressAverageGreaterThan100=NggNonMSWork_In_ProgressAverage[NggNonMSWork_In_ProgressAverage['Ageing days'] > 100]

#NggNonMSWork_In_ProgressLowAging
NggNonMSWork_In_ProgressLowLessThan10=NggNonMSWork_In_ProgressLow[NggNonMSWork_In_ProgressLow['Ageing days'] <=10]
NggNonMSWork_In_ProgressLowGreaterThan10=NggNonMSWork_In_ProgressLow[(NggNonMSWork_In_ProgressLow['Ageing days'] > 10)&(NggNonMSWork_In_ProgressLow['Ageing days'] <=50)]
NggNonMSWork_In_ProgressLowGreaterThan50=NggNonMSWork_In_ProgressLow[(NggNonMSWork_In_ProgressLow['Ageing days'] > 50)&(NggNonMSWork_In_ProgressLow['Ageing days'] <=100)]
NggNonMSWork_In_ProgressLowGreaterThan100=NggNonMSWork_In_ProgressLow[NggNonMSWork_In_ProgressLow['Ageing days'] > 100]


rowcount={'Data':['NggMSOpenCriticalLessThan10','NggMSOpenCriticalGreaterThan10','NggMSOpenCriticalGreaterThan50','NggMSOpenCriticalGreaterThan100','NggMSOpenHighLessThan10','NggMSOpenHighGreaterThan10','NggMSOpenHighGreaterThan50','NggMSOpenHighGreaterThan100','NggMSOpenAverageLessThan10','NggMSOpenAverageGreaterThan10','NggMSOpenAverageGreaterThan50','NggMSOpenAverageGreaterThan100','NggMSOpenLowLessThan10','NggMSOpenLowGreaterThan10','NggMSOpenLowGreaterThan50','NggMSOpenLowGreaterThan100','NggMSPending_ClientCriticalLessThan10','NggMSPending_ClientCriticalGreaterThan10','NggMSPending_ClientCriticalGreaterThan50','NggMSPending_ClientCriticalGreaterThan100','NggMSPending_ClientHighLessThan10','NggMSPending_ClientHighGreaterThan10','NggMSPending_ClientHighGreaterThan50','NggMSPending_ClientHighGreaterThan100','NggMSPending_ClientAverageLessThan10','NggMSPending_ClientAverageGreaterThan10','NggMSPending_ClientAverageGreaterThan50','NggMSPending_ClientAverageGreaterThan100','NggMSPending_ClientLowLessThan10','NggMSPending_ClientLowGreaterThan10','NggMSPending_ClientLowGreaterThan50','NggMSPending_ClientLowGreaterThan100','NggMSPending_Emergency_ClientCriticalLessThan10','NggMSPending_Emergency_ClientCriticalGreaterThan10','NggMSPending_Emergency_ClientCriticalGreaterThan50','NggMSPending_Emergency_ClientCriticalGreaterThan100','NggMSPending_Emergency_ClientHighLessThan10','NggMSPending_Emergency_ClientHighGreaterThan10','NggMSPending_Emergency_ClientHighGreaterThan50','NggMSPending_Emergency_ClientHighGreaterThan100','NggMSPending_Emergency_ClientAverageLessThan10','NggMSPending_Emergency_ClientAverageGreaterThan10','NggMSPending_Emergency_ClientAverageGreaterThan50','NggMSPending_Emergency_ClientAverageGreaterThan100','NggMSPending_Emergency_ClientLowLessThan10','NggMSPending_Emergency_ClientLowGreaterThan10','NggMSPending_Emergency_ClientLowGreaterThan50','NggMSPending_Emergency_ClientLowGreaterThan100','NggMSPending_VendorCriticalLessThan10','NggMSPending_VendorCriticalGreaterThan10','NggMSPending_VendorCriticalGreaterThan50','NggMSPending_VendorCriticalGreaterThan100','NggMSPending_VendorHighLessThan10','NggMSPending_VendorHighGreaterThan10','NggMSPending_VendorHighGreaterThan50','NggMSPending_VendorHighGreaterThan100','NggMSPending_VendorAverageLessThan10','NggMSPending_VendorAverageGreaterThan10','NggMSPending_VendorAverageGreaterThan50','NggMSPending_VendorAverageGreaterThan100','NggMSPending_VendorLowLessThan10','NggMSPending_VendorLowGreaterThan10','NggMSPending_VendorLowGreaterThan50','NggMSPending_VendorLowGreaterThan100','NggMSWork_In_ProgressCriticalLessThan10','NggMSWork_In_ProgressCriticalGreaterThan10','NggMSWork_In_ProgressCriticalGreaterThan50','NggMSWork_In_ProgressCriticalGreaterThan100','NggMSWork_In_ProgressHighLessThan10','NggMSWork_In_ProgressHighGreaterThan10','NggMSWork_In_ProgressHighGreaterThan50','NggMSWork_In_ProgressHighGreaterThan100','NggMSWork_In_ProgressAverageLessThan10','NggMSWork_In_ProgressAverageGreaterThan10','NggMSWork_In_ProgressAverageGreaterThan50','NggMSWork_In_ProgressAverageGreaterThan100','NggMSWork_In_ProgressLowLessThan10','NggMSWork_In_ProgressLowGreaterThan10','NggMSWork_In_ProgressLowGreaterThan50','NggMSWork_In_ProgressLowGreaterThan100','NggNonMSOpenCriticalLessThan10','NggNonMSOpenCriticalGreaterThan10','NggNonMSOpenCriticalGreaterThan50','NggNonMSOpenCriticalGreaterThan100','NggNonMSOpenHighLessThan10','NggNonMSOpenHighGreaterThan10','NggNonMSOpenHighGreaterThan50','NggNonMSOpenHighGreaterThan100','NggNonMSOpenAverageLessThan10','NggNonMSOpenAverageGreaterThan10','NggNonMSOpenAverageGreaterThan50','NggNonMSOpenAverageGreaterThan100','NggNonMSOpenLowLessThan10','NggNonMSOpenLowGreaterThan10','NggNonMSOpenLowGreaterThan50','NggNonMSOpenLowGreaterThan100','NggNonMSPending_ClientCriticalLessThan10','NggNonMSPending_ClientCriticalGreaterThan10','NggNonMSPending_ClientCriticalGreaterThan50','NggNonMSPending_ClientCriticalGreaterThan100','NggNonMSPending_ClientHighLessThan10','NggNonMSPending_ClientHighGreaterThan10','NggNonMSPending_ClientHighGreaterThan50','NggNonMSPending_ClientHighGreaterThan100','NggNonMSPending_ClientAverageLessThan10','NggNonMSPending_ClientAverageGreaterThan10','NggNonMSPending_ClientAverageGreaterThan50','NggNonMSPending_ClientAverageGreaterThan100','NggNonMSPending_ClientLowLessThan10','NggNonMSPending_ClientLowGreaterThan10','NggNonMSPending_ClientLowGreaterThan50','NggNonMSPending_ClientLowGreaterThan100','NggNonMSPending_Emergency_ClientCriticalLessThan10','NggNonMSPending_Emergency_ClientCriticalGreaterThan10','NggNonMSPending_Emergency_ClientCriticalGreaterThan50','NggNonMSPending_Emergency_ClientCriticalGreaterThan100','NggNonMSPending_Emergency_ClientHighLessThan10','NggNonMSPending_Emergency_ClientHighGreaterThan10','NggNonMSPending_Emergency_ClientHighGreaterThan50','NggNonMSPending_Emergency_ClientHighGreaterThan100','NggNonMSPending_Emergency_ClientAverageLessThan10','NggNonMSPending_Emergency_ClientAverageGreaterThan10','NggNonMSPending_Emergency_ClientAverageGreaterThan50','NggNonMSPending_Emergency_ClientAverageGreaterThan100','NggNonMSPending_Emergency_ClientLowLessThan10','NggNonMSPending_Emergency_ClientLowGreaterThan10','NggNonMSPending_Emergency_ClientLowGreaterThan50','NggNonMSPending_Emergency_ClientLowGreaterThan100','NggNonMSPending_VendorCriticalLessThan10','NggNonMSPending_VendorCriticalGreaterThan10','NggNonMSPending_VendorCriticalGreaterThan50','NggNonMSPending_VendorCriticalGreaterThan100','NggNonMSPending_VendorHighLessThan10','NggNonMSPending_VendorHighGreaterThan10','NggNonMSPending_VendorHighGreaterThan50','NggNonMSPending_VendorHighGreaterThan100','NggNonMSPending_VendorAverageLessThan10','NggNonMSPending_VendorAverageGreaterThan10','NggNonMSPending_VendorAverageGreaterThan50','NggNonMSPending_VendorAverageGreaterThan100','NggNonMSPending_VendorLowLessThan10','NggNonMSPending_VendorLowGreaterThan10','NggNonMSPending_VendorLowGreaterThan50','NggNonMSPending_VendorLowGreaterThan100','NggNonMSWork_In_ProgressCriticalLessThan10','NggNonMSWork_In_ProgressCriticalGreaterThan10','NggNonMSWork_In_ProgressCriticalGreaterThan50','NggNonMSWork_In_ProgressCriticalGreaterThan100','NggNonMSWork_In_ProgressHighLessThan10','NggNonMSWork_In_ProgressHighGreaterThan10','NggNonMSWork_In_ProgressHighGreaterThan50','NggNonMSWork_In_ProgressHighGreaterThan100','NggNonMSWork_In_ProgressAverageLessThan10','NggNonMSWork_In_ProgressAverageGreaterThan10','NggNonMSWork_In_ProgressAverageGreaterThan50','NggNonMSWork_In_ProgressAverageGreaterThan100','NggNonMSWork_In_ProgressLowLessThan10','NggNonMSWork_In_ProgressLowGreaterThan10','NggNonMSWork_In_ProgressLowGreaterThan50','NggNonMSWork_In_ProgressLowGreaterThan100'
],'RowCount':[NggMSOpenCriticalLessThan10.shape[0],NggMSOpenCriticalGreaterThan10.shape[0],NggMSOpenCriticalGreaterThan50.shape[0],NggMSOpenCriticalGreaterThan100.shape[0],NggMSOpenHighLessThan10.shape[0],NggMSOpenHighGreaterThan10.shape[0],NggMSOpenHighGreaterThan50.shape[0],NggMSOpenHighGreaterThan100.shape[0],NggMSOpenAverageLessThan10.shape[0],NggMSOpenAverageGreaterThan10.shape[0],NggMSOpenAverageGreaterThan50.shape[0],NggMSOpenAverageGreaterThan100.shape[0],NggMSOpenLowLessThan10.shape[0],NggMSOpenLowGreaterThan10.shape[0],NggMSOpenLowGreaterThan50.shape[0],NggMSOpenLowGreaterThan100.shape[0],NggMSPending_ClientCriticalLessThan10.shape[0],NggMSPending_ClientCriticalGreaterThan10.shape[0],NggMSPending_ClientCriticalGreaterThan50.shape[0],NggMSPending_ClientCriticalGreaterThan100.shape[0],NggMSPending_ClientHighLessThan10.shape[0],NggMSPending_ClientHighGreaterThan10.shape[0],NggMSPending_ClientHighGreaterThan50.shape[0],NggMSPending_ClientHighGreaterThan100.shape[0],NggMSPending_ClientAverageLessThan10.shape[0],NggMSPending_ClientAverageGreaterThan10.shape[0],NggMSPending_ClientAverageGreaterThan50.shape[0],NggMSPending_ClientAverageGreaterThan100.shape[0],NggMSPending_ClientLowLessThan10.shape[0],NggMSPending_ClientLowGreaterThan10.shape[0],NggMSPending_ClientLowGreaterThan50.shape[0],NggMSPending_ClientLowGreaterThan100.shape[0],NggMSPending_Emergency_ClientCriticalLessThan10.shape[0],NggMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],NggMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],NggMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],NggMSPending_Emergency_ClientHighLessThan10.shape[0],NggMSPending_Emergency_ClientHighGreaterThan10.shape[0],NggMSPending_Emergency_ClientHighGreaterThan50.shape[0],NggMSPending_Emergency_ClientHighGreaterThan100.shape[0],NggMSPending_Emergency_ClientAverageLessThan10.shape[0],NggMSPending_Emergency_ClientAverageGreaterThan10.shape[0],NggMSPending_Emergency_ClientAverageGreaterThan50.shape[0],NggMSPending_Emergency_ClientAverageGreaterThan100.shape[0],NggMSPending_Emergency_ClientLowLessThan10.shape[0],NggMSPending_Emergency_ClientLowGreaterThan10.shape[0],NggMSPending_Emergency_ClientLowGreaterThan50.shape[0],NggMSPending_Emergency_ClientLowGreaterThan100.shape[0],NggMSPending_VendorCriticalLessThan10.shape[0],NggMSPending_VendorCriticalGreaterThan10.shape[0],NggMSPending_VendorCriticalGreaterThan50.shape[0],NggMSPending_VendorCriticalGreaterThan100.shape[0],NggMSPending_VendorHighLessThan10.shape[0],NggMSPending_VendorHighGreaterThan10.shape[0],NggMSPending_VendorHighGreaterThan50.shape[0],NggMSPending_VendorHighGreaterThan100.shape[0],NggMSPending_VendorAverageLessThan10.shape[0],NggMSPending_VendorAverageGreaterThan10.shape[0],NggMSPending_VendorAverageGreaterThan50.shape[0],NggMSPending_VendorAverageGreaterThan100.shape[0],NggMSPending_VendorLowLessThan10.shape[0],NggMSPending_VendorLowGreaterThan10.shape[0],NggMSPending_VendorLowGreaterThan50.shape[0],NggMSPending_VendorLowGreaterThan100.shape[0],NggMSWork_In_ProgressCriticalLessThan10.shape[0],NggMSWork_In_ProgressCriticalGreaterThan10.shape[0],NggMSWork_In_ProgressCriticalGreaterThan50.shape[0],NggMSWork_In_ProgressCriticalGreaterThan100.shape[0],NggMSWork_In_ProgressHighLessThan10.shape[0],NggMSWork_In_ProgressHighGreaterThan10.shape[0],NggMSWork_In_ProgressHighGreaterThan50.shape[0],NggMSWork_In_ProgressHighGreaterThan100.shape[0],NggMSWork_In_ProgressAverageLessThan10.shape[0],NggMSWork_In_ProgressAverageGreaterThan10.shape[0],NggMSWork_In_ProgressAverageGreaterThan50.shape[0],NggMSWork_In_ProgressAverageGreaterThan100.shape[0],NggMSWork_In_ProgressLowLessThan10.shape[0],NggMSWork_In_ProgressLowGreaterThan10.shape[0],NggMSWork_In_ProgressLowGreaterThan50.shape[0],NggMSWork_In_ProgressLowGreaterThan100.shape[0],NggNonMSOpenCriticalLessThan10.shape[0],NggNonMSOpenCriticalGreaterThan10.shape[0],NggNonMSOpenCriticalGreaterThan50.shape[0],NggNonMSOpenCriticalGreaterThan100.shape[0],NggNonMSOpenHighLessThan10.shape[0],NggNonMSOpenHighGreaterThan10.shape[0],NggNonMSOpenHighGreaterThan50.shape[0],NggNonMSOpenHighGreaterThan100.shape[0],NggNonMSOpenAverageLessThan10.shape[0],NggNonMSOpenAverageGreaterThan10.shape[0],NggNonMSOpenAverageGreaterThan50.shape[0],NggNonMSOpenAverageGreaterThan100.shape[0],NggNonMSOpenLowLessThan10.shape[0],NggNonMSOpenLowGreaterThan10.shape[0],NggNonMSOpenLowGreaterThan50.shape[0],NggNonMSOpenLowGreaterThan100.shape[0],NggNonMSPending_ClientCriticalLessThan10.shape[0],NggNonMSPending_ClientCriticalGreaterThan10.shape[0],NggNonMSPending_ClientCriticalGreaterThan50.shape[0],NggNonMSPending_ClientCriticalGreaterThan100.shape[0],NggNonMSPending_ClientHighLessThan10.shape[0],NggNonMSPending_ClientHighGreaterThan10.shape[0],NggNonMSPending_ClientHighGreaterThan50.shape[0],NggNonMSPending_ClientHighGreaterThan100.shape[0],NggNonMSPending_ClientAverageLessThan10.shape[0],NggNonMSPending_ClientAverageGreaterThan10.shape[0],NggNonMSPending_ClientAverageGreaterThan50.shape[0],NggNonMSPending_ClientAverageGreaterThan100.shape[0],NggNonMSPending_ClientLowLessThan10.shape[0],NggNonMSPending_ClientLowGreaterThan10.shape[0],NggNonMSPending_ClientLowGreaterThan50.shape[0],NggNonMSPending_ClientLowGreaterThan100.shape[0],NggNonMSPending_Emergency_ClientCriticalLessThan10.shape[0],NggNonMSPending_Emergency_ClientCriticalGreaterThan10.shape[0],NggNonMSPending_Emergency_ClientCriticalGreaterThan50.shape[0],NggNonMSPending_Emergency_ClientCriticalGreaterThan100.shape[0],NggNonMSPending_Emergency_ClientHighLessThan10.shape[0],NggNonMSPending_Emergency_ClientHighGreaterThan10.shape[0],NggNonMSPending_Emergency_ClientHighGreaterThan50.shape[0],NggNonMSPending_Emergency_ClientHighGreaterThan100.shape[0],NggNonMSPending_Emergency_ClientAverageLessThan10.shape[0],NggNonMSPending_Emergency_ClientAverageGreaterThan10.shape[0],NggNonMSPending_Emergency_ClientAverageGreaterThan50.shape[0],NggNonMSPending_Emergency_ClientAverageGreaterThan100.shape[0],NggNonMSPending_Emergency_ClientLowLessThan10.shape[0],NggNonMSPending_Emergency_ClientLowGreaterThan10.shape[0],NggNonMSPending_Emergency_ClientLowGreaterThan50.shape[0],NggNonMSPending_Emergency_ClientLowGreaterThan100.shape[0],NggNonMSPending_VendorCriticalLessThan10.shape[0],NggNonMSPending_VendorCriticalGreaterThan10.shape[0],NggNonMSPending_VendorCriticalGreaterThan50.shape[0],NggNonMSPending_VendorCriticalGreaterThan100.shape[0],NggNonMSPending_VendorHighLessThan10.shape[0],NggNonMSPending_VendorHighGreaterThan10.shape[0],NggNonMSPending_VendorHighGreaterThan50.shape[0],NggNonMSPending_VendorHighGreaterThan100.shape[0],NggNonMSPending_VendorAverageLessThan10.shape[0],NggNonMSPending_VendorAverageGreaterThan10.shape[0],NggNonMSPending_VendorAverageGreaterThan50.shape[0],NggNonMSPending_VendorAverageGreaterThan100.shape[0],NggNonMSPending_VendorLowLessThan10.shape[0],NggNonMSPending_VendorLowGreaterThan10.shape[0],NggNonMSPending_VendorLowGreaterThan50.shape[0],NggNonMSPending_VendorLowGreaterThan100.shape[0],NggNonMSWork_In_ProgressCriticalLessThan10.shape[0],NggNonMSWork_In_ProgressCriticalGreaterThan10.shape[0],NggNonMSWork_In_ProgressCriticalGreaterThan50.shape[0],NggNonMSWork_In_ProgressCriticalGreaterThan100.shape[0],NggNonMSWork_In_ProgressHighLessThan10.shape[0],NggNonMSWork_In_ProgressHighGreaterThan10.shape[0],NggNonMSWork_In_ProgressHighGreaterThan50.shape[0],NggNonMSWork_In_ProgressHighGreaterThan100.shape[0],NggNonMSWork_In_ProgressAverageLessThan10.shape[0],NggNonMSWork_In_ProgressAverageGreaterThan10.shape[0],NggNonMSWork_In_ProgressAverageGreaterThan50.shape[0],NggNonMSWork_In_ProgressAverageGreaterThan100.shape[0],NggNonMSWork_In_ProgressLowLessThan10.shape[0],NggNonMSWork_In_ProgressLowGreaterThan10.shape[0],NggNonMSWork_In_ProgressLowGreaterThan50.shape[0],NggNonMSWork_In_ProgressLowGreaterThan100.shape[0]
]}
rcdf=pd.DataFrame(rowcount)
rcdf.to_excel(rowpath,index=False)
print("success")

#rowcount
#NggMSOpenCriticalLessThan10rc=NggMSOpenCriticalLessThan10.shape[0]
##(NggMSWork_In_ProgressHighLessThan10)
##(MsNgg.shape[0])
##(NggMSOpen)
##(NggMSOpenCritical)













#writer = pd.ExcelWriter(excelPath.replace(".xlsx","1.xlsx"), engine='xlsxwriter')

#ngg1.to_excel(writer, sheet_name=ng,index=False)
#ediSupport1.to_excel(writer, sheet_name=es,index=False)
#ediAnalysts1.to_excel(writer, sheet_name=ea,index=False)
#webMethods1.to_excel(writer, sheet_name=we,index=False)

#writer.save()
##("success")
