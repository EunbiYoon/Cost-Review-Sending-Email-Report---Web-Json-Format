import pandas as pd
import openpyxl 
import sys

#automatic working from Adata
# sys.path.insert(0,'C:/Users/RnD Workstation/Documents/CostReview/PGM/Report_Email')
# from Adata import TBI_P_2 as read_value

#manually read excel 
read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0616/Cost Review_0616.xlsx", sheet_name="DR_PAC_Item")
read_excel=read_excel.drop('Unnamed: 0', axis=1)

#round 2
for i in range(len(read_excel)):
    a=read_excel.at[i,'VI']
    b=read_excel.at[i,'VI.1']
    read_excel.at[i,'VI']=a.round(2)
    read_excel.at[i,'VI.1']=b.round(2)

# read_increase=read_excel[["Increase","VI","Date","Reason"]]
read_increase=read_excel[["Increase","VI","Date"]]
read_increase=read_increase.dropna()
increase_total=read_increase["VI"].sum().round(2)
read_increase.index=range(1,len(read_increase)+1)
read_increase.at["Total","VI"]=increase_total
read_increase=read_increase.fillna('')

# read_decrease=read_excel[["Decrease","VI.1","Date.1","Reason.1"]]
read_decrease=read_excel[["Decrease","VI.1","Date.1"]]
read_decrease=read_decrease.rename(columns={"VI.1": "VI", "Date.1": "Date","Reason.1":"Reason"})
read_decrease=read_decrease.dropna()
decrease_total=read_decrease["VI"].sum().round(2)
read_decrease.index=range(1,len(read_decrease)+1)
read_decrease.at["Total","VI"]=decrease_total
read_decrease=read_decrease.fillna('')


########### increase_item
print('"Increase":{')

#column_list
column_list=list(read_increase.columns)
column_list.insert(0,"")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""').replace("blank","").replace(".1","").replace("Unnamed: 0","")
print('  '+column_json+",")

#row_list
read_value=read_increase.T
read_value.reset_index(inplace=True,drop=True)
read_value=read_value.T
read_value.reset_index(inplace=True)

#row_list - index
bb=pd.DataFrame()
for i in range(len(read_value.index)):
    bb.at[0,i]='"index":"'+str(read_value.at[i,"index"])+'",'

print('  '+'"rows":[')
#row_list - values
for i in range(len(read_value.index)):
    if i==len(read_value.index)-1:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"}"
        print('  '+AA)
    else:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"},"
        print('  '+AA)
print(']},')

########### decrease_item
print('"Decrease":{')
#column_list
column_list=list(read_decrease.columns)
column_list.insert(0,"")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""').replace("blank","").replace(".1","").replace("Unnamed: 0","")
print('  '+column_json+",")

#row_list
read_value=read_decrease.T
read_value.reset_index(inplace=True,drop=True)
read_value=read_value.T
read_value.reset_index(inplace=True)

#row_list - index
bb=pd.DataFrame()
for i in range(len(read_value.index)):
    bb.at[0,i]='"index":"'+str(read_value.at[i,"index"])+'",'

print('  '+'"rows":[')
#row_list - values
for i in range(len(read_value.index)):
    if i==len(read_value.index)-1:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"}"
        print('  '+AA)
    else:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"},"
        print('  '+AA)
print(']}')