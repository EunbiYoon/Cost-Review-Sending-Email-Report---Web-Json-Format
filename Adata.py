import pandas as pd
import openpyxl
import datetime
import numpy as np

#This weeknum
this_week="23.06 W5"
next_month1="23.07"
next_month2="23.08"

#lastweek
last_result="Cost Review_last"

#add the week
# next_week1="23.06 W5"

#read data - last and original
data_last="0623"
data_this="0628"

############################ Trend Table ############################  
#read original data
F_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="FL_BPA")
F_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="FL_PAC")
T_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="TL_BPA")
T_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="TL_PAC")
D_original_BPAE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="DR_BPA")
D_original_PACE=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/"+last_result+".xlsx", sheet_name="DR_PAC")

#read new data
bpa_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="BPA Entity")
pac_entity=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="PAC Entity")

#remain required  new data
BPAE=bpa_entity[["Model.Suffix","Net RMC (USD)"]]
# BPAE=BPAE.drop([0])
BPAE.reset_index(inplace=True, drop=True)
for i in range(len(BPAE)):
    a=BPAE.at[i,"Net RMC (USD)"]
    BPAE.at[i,"Net RMC (USD)"]=round(a,1)


PACE=pac_entity[["Model.Suffix","Net RMC (USD)"]]
# PACE=PACE.drop([0])
PACE.reset_index(inplace=True, drop=True)
PACE["Net RMC (USD)"].round(1)
for i in range(len(PACE)):
    a=PACE.at[i,"Net RMC (USD)"]
    PACE.at[i,"Net RMC (USD)"]=round(a,1)


#Matching column with original data
BPAE=BPAE.rename(columns={"Model.Suffix":"Tool","Net RMC (USD)":this_week})
PACE=PACE.rename(columns={"Model.Suffix":"Tool","Net RMC (USD)":this_week})

#merge two file with Model Suffix
F_BPAE_Merge=pd.merge(F_original_BPAE,BPAE,how="inner",on="Tool")
F_PACE_Merge=pd.merge(F_original_PACE,PACE,how="inner",on="Tool")
T_BPAE_Merge=pd.merge(T_original_BPAE,BPAE,how="inner",on="Tool")
T_PACE_Merge=pd.merge(T_original_PACE,PACE,how="inner",on="Tool")
D_BPAE_Merge=pd.merge(D_original_BPAE,BPAE,how="inner",on="Tool")
D_PACE_Merge=pd.merge(D_original_PACE,PACE,how="inner",on="Tool")

#drop unamed:0
F_BPAE_Merge=F_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
F_PACE_Merge=F_PACE_Merge.drop(['Unnamed: 0'],axis=1)
T_BPAE_Merge=T_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
T_PACE_Merge=T_PACE_Merge.drop(['Unnamed: 0'],axis=1)
D_BPAE_Merge=D_BPAE_Merge.drop(['Unnamed: 0'],axis=1)
D_PACE_Merge=D_PACE_Merge.drop(['Unnamed: 0'],axis=1)



# add the expected value
# F_BPAE_Merge[next_week1]=round(F_BPAE_Merge[this_week]-0.1,1)
# F_BPAE_Merge[next_week2]=round(F_BPAE_Merge[this_week]-0.2,1)
# F_BPAE_Merge[next_week3]=round(F_BPAE_Merge[this_week]-0.3,1)
# F_BPAE_Merge[next_week4]=round(F_BPAE_Merge[this_week]-0.4,1)
F_BPAE_Merge[next_month1]=round(F_BPAE_Merge[this_week]-0.5,1)
F_BPAE_Merge[next_month2]=round(F_BPAE_Merge[this_week]-1,1)

# F_PACE_Merge[next_week1]=round(F_PACE_Merge[this_week]-0.1,1)
# F_PACE_Merge[next_week2]=round(F_PACE_Merge[this_week]-0.2,1)
# F_PACE_Merge[next_week3]=round(F_PACE_Merge[this_week]-0.3,1)
# F_PACE_Merge[next_week4]=round(F_PACE_Merge[this_week]-0.4,1)
F_PACE_Merge[next_month1]=round(F_PACE_Merge[this_week]-0.5,1)
F_PACE_Merge[next_month2]=round(F_PACE_Merge[this_week]-1,1)

# T_BPAE_Merge[next_week1]=round(T_BPAE_Merge[this_week]-0.1,1)
# T_BPAE_Merge[next_week2]=round(T_BPAE_Merge[this_week]-0.2,1)
# T_BPAE_Merge[next_week3]=round(T_BPAE_Merge[this_week]-0.3,1)
# T_BPAE_Merge[next_week4]=round(T_BPAE_Merge[this_week]-0.4,1)
T_BPAE_Merge[next_month1]=round(T_BPAE_Merge[this_week]-0.5,1)
T_BPAE_Merge[next_month2]=round(T_BPAE_Merge[this_week]-1,1)
print(F_PACE_Merge)

# T_PACE_Merge[next_week1]=round(T_PACE_Merge[this_week]-0.1,1)
# T_PACE_Merge[next_week2]=round(T_PACE_Merge[this_week]-0.2,1)
# T_PACE_Merge[next_week3]=round(T_PACE_Merge[this_week]-0.3,1)
# T_PACE_Merge[next_week4]=round(T_PACE_Merge[this_week]-0.4,1)
T_PACE_Merge[next_month1]=round(T_PACE_Merge[this_week]-0.5,1)
T_PACE_Merge[next_month2]=round(T_PACE_Merge[this_week]-1,1)

# D_BPAE_Merge[next_week1]=round(D_BPAE_Merge[this_week]-0.1,1)
# D_BPAE_Merge[next_week2]=round(D_BPAE_Merge[this_week]-0.2,1)
# D_BPAE_Merge[next_week3]=round(D_BPAE_Merge[this_week]-0.3,1)
# D_BPAE_Merge[next_week4]=round(D_BPAE_Merge[this_week]-0.4,1)
D_BPAE_Merge[next_month1]=round(D_BPAE_Merge[this_week]-0.5,1)
D_BPAE_Merge[next_month2]=round(D_BPAE_Merge[this_week]-1,1)

# D_PACE_Merge[next_week1]=round(D_PACE_Merge[this_week]-0.1,1)
# D_PACE_Merge[next_week2]=round(D_PACE_Merge[this_week]-0.2,1)
# D_PACE_Merge[next_week3]=round(D_PACE_Merge[this_week]-0.3,1)
# D_PACE_Merge[next_week4]=round(D_PACE_Merge[this_week]-0.4,1)
D_PACE_Merge[next_month1]=round(D_PACE_Merge[this_week]-0.5,1)
D_PACE_Merge[next_month2]=round(D_PACE_Merge[this_week]-1,1)


# change index
F_BPAE_Merge.index=range(1,len(F_BPAE_Merge)+1)
F_PACE_Merge.index=range(1,len(F_PACE_Merge)+1)
T_BPAE_Merge.index=range(1,len(T_BPAE_Merge)+1)
T_PACE_Merge.index=range(1,len(T_PACE_Merge)+1)
D_BPAE_Merge.index=range(1,len(D_BPAE_Merge)+1)
D_PACE_Merge.index=range(1,len(D_PACE_Merge)+1)



############################ Item Table ############################  
#data read
F_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="FL_BPA_Item")
D_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="DR_BPA_Item")
T_BPA_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="TL_BPA_Item")

F_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="FL_PAC_Item")
D_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="DR_PAC_Item")
T_PAC_I=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/data.xlsx", sheet_name="TL_PAC_Item")

#data required
F_BPA_I=F_BPA_I.drop([0],axis=0)
F_BPA_I=F_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
F_BPA_I=F_BPA_I.dropna()
F_BPA_I=F_BPA_I.drop(['Gap'],axis=1)
F_BPA_I=F_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
F_BPA_I.reset_index(drop=True, inplace=True)

D_BPA_I=D_BPA_I.drop([0],axis=0)
D_BPA_I=D_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
D_BPA_I=D_BPA_I.dropna()
D_BPA_I=D_BPA_I.drop(['Gap'],axis=1)
D_BPA_I=D_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
D_BPA_I.reset_index(drop=True, inplace=True)

T_BPA_I=T_BPA_I.drop([0],axis=0)
T_BPA_I=T_BPA_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
T_BPA_I=T_BPA_I.dropna()
T_BPA_I=T_BPA_I.drop(['Gap'],axis=1)
T_BPA_I=T_BPA_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
T_BPA_I.reset_index(drop=True, inplace=True)

F_PAC_I=F_PAC_I.drop([0],axis=0)
F_PAC_I=F_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
F_PAC_I=F_PAC_I.dropna()
F_PAC_I=F_PAC_I.drop(['Gap'],axis=1)
F_PAC_I=F_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
F_PAC_I.reset_index(drop=True, inplace=True)

D_PAC_I=D_PAC_I.drop([0],axis=0)
D_PAC_I=D_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
D_PAC_I=D_PAC_I.dropna()
D_PAC_I=D_PAC_I.drop(['Gap'],axis=1)
D_PAC_I=D_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
D_PAC_I.reset_index(drop=True, inplace=True)

T_PAC_I=T_PAC_I.drop([0],axis=0)
T_PAC_I=T_PAC_I[['[Left] Model/PartNo.3','[Left] Model/PartNo.4','Gap','Gap.1']]
T_PAC_I=T_PAC_I.dropna()
T_PAC_I=T_PAC_I.drop(['Gap'],axis=1)
T_PAC_I=T_PAC_I.rename(columns={"[Left] Model/PartNo.3": "Part No", "[Left] Model/PartNo.4": "Description","Gap.1":"VI"})
T_PAC_I.reset_index(drop=True, inplace=True)


# part number same merge
FBI=pd.DataFrame(F_BPA_I.groupby(['Part No','Description']).sum())
FBI.reset_index(inplace=True)
TBI=pd.DataFrame(T_BPA_I.groupby(['Part No','Description']).sum())
TBI.reset_index(inplace=True)
DBI=pd.DataFrame(D_BPA_I.groupby(['Part No','Description']).sum())
DBI.reset_index(inplace=True)
FPI=pd.DataFrame(F_PAC_I.groupby(['Part No','Description']).sum())
FPI.reset_index(inplace=True)
TPI=pd.DataFrame(T_PAC_I.groupby(['Part No','Description']).sum())
TPI.reset_index(inplace=True)
DPI=pd.DataFrame(D_PAC_I.groupby(['Part No','Description']).sum())
DPI.reset_index(inplace=True)

#sort top 5 (increase:2, decrease :3)
FBI=FBI.sort_values(by='VI',ascending=True)
FBI_L=FBI[:3]
FBI=FBI.sort_values(by='VI',ascending=False)
FBI_H=FBI[:2]

TBI=TBI.sort_values(by='VI',ascending=True)
TBI_L=TBI[:3]
TBI=TBI.sort_values(by='VI',ascending=False)
TBI_H=TBI[:2]

DBI=DBI.sort_values(by='VI',ascending=True)
DBI_L=DBI[:3]
DBI=DBI.sort_values(by='VI',ascending=False)
DBI_H=DBI[:2]

FPI=FPI.sort_values(by='VI',ascending=True)
FPI_L=FPI[:3]
FPI=FPI.sort_values(by='VI',ascending=False)
FPI_H=FPI[:2]

TPI=TPI.sort_values(by='VI',ascending=True)
TPI_L=TPI[:3]
TPI=TPI.sort_values(by='VI',ascending=False)
TPI_H=TPI[:2]

DPI=DPI.sort_values(by='VI',ascending=True)
DPI_L=DPI[:3]
DPI=DPI.sort_values(by='VI',ascending=False)
DPI_H=DPI[:2]


#high value -> higher than 0
FBI_H.reset_index(inplace=True, drop=True)
for i in range(len(FBI_H)):
    condition=FBI_H.at[i,"VI"]
    if condition<0:
        FBI_H=FBI_H.drop([i],axis=0)

TBI_H.reset_index(inplace=True, drop=True)
for i in range(len(FBI_H)):
    condition=TBI_H.at[i,"VI"]
    if condition<0:
        TBI_H=TBI_H.drop([i],axis=0)

DBI_H.reset_index(inplace=True, drop=True)
for i in range(len(DBI_H)):
    condition=DBI_H.at[i,"VI"]
    if condition<0:
        DBI_H=DBI_H.drop([i],axis=0)

FPI_H.reset_index(inplace=True, drop=True)
for i in range(len(FPI_H)):
    condition=FPI_H.at[i,"VI"]
    if condition<0:
        FPI_H=FPI_H.drop([i],axis=0)

TPI_H.reset_index(inplace=True, drop=True)
for i in range(len(TPI_H)):
    condition=TPI_H.at[i,"VI"]
    if condition<0:
        TPI_H=TPI_H.drop([i],axis=0)

DPI_H.reset_index(inplace=True, drop=True)
for i in range(len(DPI_H)):
    condition=DPI_H.at[i,"VI"]
    if condition<0:
        DPI_H=DPI_H.drop([i],axis=0)

#lower value -> lower than 0
FBI_L.reset_index(inplace=True, drop=True)
for i in range(len(FBI_L)):
    condition=FBI_L.at[i,"VI"]
    if condition>0:
        FBI_L=FBI_L.drop([i],axis=0)

TBI_L.reset_index(inplace=True, drop=True)
for i in range(len(TBI_L)):
    condition=TBI_L.at[i,"VI"]
    if condition>0:
        TBI_L=TBI_L.drop([i],axis=0)

DBI_L.reset_index(inplace=True, drop=True)
for i in range(len(DBI_L)):
    condition=DBI_L.at[i,"VI"]
    if condition>0:
        DBI_L=DBI_L.drop([i],axis=0)

FPI_L.reset_index(inplace=True, drop=True)
for i in range(len(FPI_L)):
    condition=FPI_L.at[i,"VI"]
    if condition>0:
        FPI_L=FPI_L.drop([i],axis=0)

TPI_L.reset_index(inplace=True, drop=True)
for i in range(len(TPI_L)):
    condition=TPI_L.at[i,"VI"]
    if condition>0:
        TPI_L=TPI_L.drop([i],axis=0)

DPI_L.reset_index(inplace=True, drop=True)
for i in range(len(DPI_L)):
    condition=DPI_L.at[i,"VI"]
    if condition>0:
        DPI_L=DPI_L.drop([i],axis=0)

#reset H and L
FBI_H.reset_index(inplace=True, drop=True)
TBI_H.reset_index(inplace=True, drop=True)
DBI_H.reset_index(inplace=True, drop=True)
FPI_H.reset_index(inplace=True, drop=True)
TPI_H.reset_index(inplace=True, drop=True)
DPI_H.reset_index(inplace=True, drop=True)

FBI_L.reset_index(inplace=True, drop=True)
TBI_L.reset_index(inplace=True, drop=True)
DBI_L.reset_index(inplace=True, drop=True)
FPI_L.reset_index(inplace=True, drop=True)
TPI_L.reset_index(inplace=True, drop=True)
DPI_L.reset_index(inplace=True, drop=True)

# read previous report and merge
FBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="FL_BPA_Item")
TBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="TL_BPA_Item")
DBI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="DR_BPA_Item")
FPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="FL_PAC_Item")
TPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="TL_PAC_Item")
DPI_P=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/"+data_last+"/Cost Review_"+data_last+".xlsx", sheet_name="DR_PAC_Item")

FBI_P.index=FBI_P["Unnamed: 0"]
FBI_P=FBI_P.drop(["Unnamed: 0"], axis=1)
FBI_P=FBI_P.rename_axis(None)
FBI_P_1=FBI_P[['Increase','VI','Date']].dropna()
FBI_P_2=FBI_P[['Decrease','VI.1','Date.1']].dropna()

TBI_P.index=TBI_P["Unnamed: 0"]
TBI_P=TBI_P.drop(["Unnamed: 0"], axis=1)
TBI_P=TBI_P.rename_axis(None)
TBI_P_1=TBI_P[['Increase','VI','Date']].dropna()
TBI_P_2=TBI_P[['Decrease','VI.1','Date.1']].dropna()

DBI_P.index=DBI_P["Unnamed: 0"]
DBI_P=DBI_P.drop(["Unnamed: 0"], axis=1)
DBI_P=DBI_P.rename_axis(None)
DBI_P_1=DBI_P[['Increase','VI','Date']].dropna()
DBI_P_2=DBI_P[['Decrease','VI.1','Date.1']].dropna()

FPI_P.index=FPI_P["Unnamed: 0"]
FPI_P=FPI_P.drop(["Unnamed: 0"], axis=1)
FPI_P=FPI_P.rename_axis(None)
FPI_P_1=FPI_P[['Increase','VI','Date']].dropna()
FPI_P_2=FPI_P[['Decrease','VI.1','Date.1']].dropna()

TPI_P.index=TPI_P["Unnamed: 0"]
TPI_P=TPI_P.drop(["Unnamed: 0"], axis=1)
TPI_P=TPI_P.rename_axis(None)
TPI_P_1=TPI_P[['Increase','VI','Date']].dropna()
TPI_P_2=TPI_P[['Decrease','VI.1','Date.1']].dropna()

DPI_P.index=DPI_P["Unnamed: 0"]
DPI_P=DPI_P.drop(["Unnamed: 0"], axis=1)
DPI_P=DPI_P.rename_axis(None)
DPI_P_1=DPI_P[['Increase','VI','Date']].dropna()
DPI_P_2=DPI_P[['Decrease','VI.1','Date.1']].dropna()

# insert to old data
row_1=len(FBI_P_1)
row_2=len(FBI_P_2)
for i in range(len(FBI_H)):
    row_1=row_1+1
    FBI_P_1.at[row_1,"Increase"]=FBI_H.at[i,"Description"]
    FBI_P_1.at[row_1,"VI"]=FBI_H.at[i,"VI"]
    FBI_P_1.at[row_1,"Date"]=this_week
for i in range(len(FBI_L)):
    row_2=row_2+1
    FBI_P_2.at[row_2,"Decrease"]=FBI_L.at[i,"Description"]
    FBI_P_2.at[row_2,"VI.1"]=FBI_L.at[i,"VI"]
    FBI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(TBI_P_1)
row_2=len(TBI_P_2)
for i in range(len(TBI_H)):
    row_1=row_1+1
    TBI_P_1.at[row_1,"Increase"]=TBI_H.at[i,"Description"]
    TBI_P_1.at[row_1,"VI"]=TBI_H.at[i,"VI"]
    TBI_P_1.at[row_1,"Date"]=this_week
for i in range(len(TBI_L)):
    row_2=row_2+1
    TBI_P_2.at[row_2,"Decrease"]=TBI_L.at[i,"Description"]
    TBI_P_2.at[row_2,"VI.1"]=TBI_L.at[i,"VI"]
    TBI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(DBI_P_1)
row_2=len(DBI_P_2)
for i in range(len(DBI_H)):
    row_1=row_1+1
    DBI_P_1.at[row_1,"Increase"]=DBI_H.at[i,"Description"]
    DBI_P_1.at[row_1,"VI"]=DBI_H.at[i,"VI"]
    DBI_P_1.at[row_1,"Date"]=this_week

for i in range(len(DBI_L)):
    row_2=row_2+1
    DBI_P_2.at[row_2,"Decrease"]=DBI_L.at[i,"Description"]
    DBI_P_2.at[row_2,"VI.1"]=DBI_L.at[i,"VI"]
    DBI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(FPI_P_1)
row_2=len(FPI_P_2)
for i in range(len(FPI_H)):
    row_1=row_1+1
    FPI_P_1.at[row_1,"Increase"]=FPI_H.at[i,"Description"]
    FPI_P_1.at[row_1,"VI"]=FPI_H.at[i,"VI"]
    FPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(FPI_L)):
    row_2=row_2+1
    FPI_P_2.at[row_2,"Decrease"]=FPI_L.at[i,"Description"]
    FPI_P_2.at[row_2,"VI.1"]=FPI_L.at[i,"VI"]
    FPI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(TPI_P_1)
row_2=len(TPI_P_2)
for i in range(len(TPI_H)):
    row_1=row_1+1
    TPI_P_1.at[row_1,"Increase"]=TPI_H.at[i,"Description"]
    TPI_P_1.at[row_1,"VI"]=TPI_H.at[i,"VI"]
    TPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(TPI_L)):
    row_2=row_2+1
    TPI_P_2.at[row_2,"Decrease"]=TPI_L.at[i,"Description"]
    TPI_P_2.at[row_2,"VI.1"]=TPI_L.at[i,"VI"]
    TPI_P_2.at[row_2,"Date.1"]=this_week

row_1=len(DPI_P_1)
row_2=len(DPI_P_2)
for i in range(len(DPI_H)):
    row_1=row_1+1
    DPI_P_1.at[row_1,"Increase"]=DPI_H.at[i,"Description"]
    DPI_P_1.at[row_1,"VI"]=DPI_H.at[i,"VI"]
    DPI_P_1.at[row_1,"Date"]=this_week
for i in range(len(DPI_L)):
    row_2=row_2+1
    DPI_P_2.at[row_2,"Decrease"]=DPI_L.at[i,"Description"]
    DPI_P_2.at[row_2,"VI.1"]=DPI_L.at[i,"VI"]
    DPI_P_2.at[row_2,"Date.1"]=this_week




#Total value
FBI_P_1_sum=FBI_P_1.sum()
FBI_P_1.at["Total","VI"]=FBI_P_1_sum["VI"]
FBI_P_1=FBI_P_1.fillna("").round(2)

FBI_P_2_sum=FBI_P_2.sum()
FBI_P_2.at["Total","VI.1"]=FBI_P_2_sum["VI.1"]
FBI_P_2=FBI_P_2.fillna("").round(2)

TBI_P_1_sum=TBI_P_1.sum()
TBI_P_1.at["Total","VI"]=TBI_P_1_sum["VI"]
TBI_P_1=TBI_P_1.fillna("").round(2)

TBI_P_2_sum=TBI_P_2.sum()
TBI_P_2.at["Total","VI.1"]=TBI_P_2_sum["VI.1"]
TBI_P_2=TBI_P_2.fillna("").round(2)

DBI_P_1_sum=DBI_P_1.sum()
DBI_P_1.at["Total","VI"]=DBI_P_1_sum["VI"]
DBI_P_1=DBI_P_1.fillna("").round(2)

DBI_P_2_sum=DBI_P_2.sum()
DBI_P_2.at["Total","VI.1"]=DBI_P_2_sum["VI.1"]
DBI_P_2=DBI_P_2.fillna("").round(2)

FPI_P_1_sum=FPI_P_1.sum()
FPI_P_1.at["Total","VI"]=FPI_P_1_sum["VI"]
FPI_P_1=FPI_P_1.fillna("").round(2)

FPI_P_2_sum=FPI_P_2.sum()
FPI_P_2.at["Total","VI.1"]=FPI_P_2_sum["VI.1"]
FPI_P_2=FPI_P_2.fillna("").round(2)

TPI_P_1_sum=TPI_P_1.sum()
TPI_P_1.at["Total","VI"]=TPI_P_1_sum["VI"]
TPI_P_1=TPI_P_1.fillna("").round(2)

TPI_P_2_sum=TPI_P_2.sum()
TPI_P_2.at["Total","VI.1"]=TPI_P_2_sum["VI.1"]
TPI_P_2=TPI_P_2.fillna("").round(2)

DPI_P_1_sum=DPI_P_1.sum()
DPI_P_1.at["Total","VI"]=DPI_P_1_sum["VI"]
DPI_P_1=DPI_P_1.fillna("").round(2)

DPI_P_2_sum=DPI_P_2.sum()
DPI_P_2.at["Total","VI.1"]=DPI_P_2_sum["VI.1"]
DPI_P_2=DPI_P_2.fillna("").round(2)


#rename the decrease VI.1 Date.1
FBI_P_2=FBI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})
TBI_P_2=TBI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})
DBI_P_2=DBI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})

FPI_P_2=FPI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})
TPI_P_2=TPI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})
DPI_P_2=DPI_P_2.rename(columns={"VI.1":"VI","Date.1":"Date"})



############################ Write excel change data ############################  
#remove all total column for excel file format
FBI1=FBI_P_1.drop("Total",axis=0)
FBI2=FBI_P_2.drop("Total",axis=0)

TBI1=TBI_P_1.drop("Total",axis=0)
TBI2=TBI_P_2.drop("Total",axis=0)

DBI1=DBI_P_1.drop("Total",axis=0)
DBI2=DBI_P_2.drop("Total",axis=0)

FPI1=FPI_P_1.drop("Total",axis=0)
FPI2=FPI_P_2.drop("Total",axis=0)

TPI1=TPI_P_1.drop("Total",axis=0)
TPI2=TPI_P_2.drop("Total",axis=0)

DPI1=DPI_P_1.drop("Total",axis=0)
DPI2=DPI_P_2.drop("Total",axis=0)

#merge increase and decrease again
FBI_merge = pd.concat([FBI1, FBI2], axis=1)
TBI_merge = pd.concat([TBI1, TBI2], axis=1)
DBI_merge = pd.concat([DBI1, DBI2], axis=1)
FPI_merge = pd.concat([FPI1, FPI2], axis=1)
TPI_merge = pd.concat([TPI1, TPI2], axis=1)
DPI_merge = pd.concat([DPI1, DPI2], axis=1)

############################ Write excel ############################  
file_writer = pd.ExcelWriter("C:/Users/RnD Workstation/Documents/CostReview/"+data_this+"/Cost Review_"+data_this+".xlsx", engine="xlsxwriter")

F_BPAE_Merge.to_excel(file_writer, sheet_name="FL_BPA")
FBI_merge.to_excel(file_writer, sheet_name="FL_BPA_Item")

T_BPAE_Merge.to_excel(file_writer, sheet_name="TL_BPA")
TBI_merge.to_excel(file_writer, sheet_name="TL_BPA_Item")

D_BPAE_Merge.to_excel(file_writer, sheet_name="DR_BPA")
DBI_merge.to_excel(file_writer, sheet_name="DR_BPA_Item")

F_PACE_Merge.to_excel(file_writer, sheet_name="FL_PAC")
FPI_merge.to_excel(file_writer, sheet_name="FL_PAC_Item")

T_PACE_Merge.to_excel(file_writer, sheet_name="TL_PAC")
TPI_merge.to_excel(file_writer, sheet_name="TL_PAC_Item")

D_PACE_Merge.to_excel(file_writer, sheet_name="DR_PAC")
DPI_merge.to_excel(file_writer, sheet_name="DR_PAC_Item")

file_writer.close()


