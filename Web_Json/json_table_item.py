import pandas as pd
import openpyxl 
import sys
sys.path.insert(0,'C:/Users/RnD Workstation/Documents/CostReview/PGM/Report_Email')
from data_table import TPI_P_2 as read_value


########### increase_item
#column_list
column_list=list(read_value.columns)
column_list.insert(0,"")

# column json file format
column_json=str({"columns":column_list}).replace("{",'').replace("}",'').replace("'",'"').replace('nan','""').replace("blank","").replace(".1","").replace("Unnamed: 0","")
print("")
print(column_json+",")

#row_list
read_value=read_value.T
read_value.reset_index(inplace=True,drop=True)
read_value=read_value.T
read_value.reset_index(inplace=True)

#row_list - index
bb=pd.DataFrame()
for i in range(len(read_value.index)):
    bb.at[0,i]='"index":"'+str(read_value.at[i,"index"])+'",'

print('"rows":[')
#row_list - values
for i in range(len(read_value.index)):
    if i==len(read_value.index)-1:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"}"
        print(AA)
    else:
        aa=str(list(read_value.iloc[i][1:])).replace("'",'"').replace('nan','""')
        aA='"values":'+aa
        AA="{"+str(bb.at[0,i])+aA+"},"
        print(AA)
print(']')
print('')