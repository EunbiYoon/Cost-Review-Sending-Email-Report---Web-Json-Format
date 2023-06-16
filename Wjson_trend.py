import pandas as pd
import openpyxl 

read_excel=pd.read_excel("C:/Users/RnD Workstation/Documents/CostReview/0616/Cost Review_0616.xlsx", sheet_name="DR_PAC")

#column_list
column_list=list(read_excel.columns)
# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""')
print("")
print(column_json+",")

#row_list
read_excel=read_excel.T
read_excel.reset_index(inplace=True, drop=True)
read_excel=read_excel.T


#round 1
for i in range(len(read_excel.index)):
    for j in range(len(read_excel.columns)-3):
        a=read_excel.at[i,j+3]
        read_excel.at[i,j+3]=round(a,1)

#row_list - index
bb=pd.DataFrame()
for i in range(len(read_excel.index)):
    bb.at[0,i]='"index":"'+str(read_excel.at[i,0])+'",'

print('"rows":[')
#row_list - values
for i in range(len(read_excel.index)):
    if i==len(read_excel.index)-1:
        aa=str(list(read_excel.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"}"
        print(AA)
    else:
        aa=str(list(read_excel.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"},"
        print(AA)
print(']')