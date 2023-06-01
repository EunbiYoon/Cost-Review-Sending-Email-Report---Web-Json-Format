from Adata import F_BPAE_Merge,T_BPAE_Merge,D_BPAE_Merge,F_PACE_Merge,T_PACE_Merge,D_PACE_Merge,DBI_P_1,DBI_P_2,TBI_P_1, TBI_P_2,FBI_P_1,FBI_P_2,DPI_P_1,DPI_P_2,TPI_P_1, TPI_P_2,FPI_P_1,FPI_P_2,this_week

############################ html  ############################  
# change to html -> table & border
F_BPAE_html=F_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
T_BPAE_html=T_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
D_BPAE_html=D_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
F_PACE_html=F_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
T_PACE_html=T_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
D_PACE_html=D_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')

FBI1_html=FBI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
FBI2_html=FBI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TBI1_html=TBI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TBI2_html=TBI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DBI1_html=DBI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DBI2_html=DBI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')

FPI1_html=FPI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
FPI2_html=FPI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TPI1_html=TPI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TPI2_html=TPI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DPI1_html=DPI_P_1.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DPI2_html=DPI_P_2.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')



# text align center & column color
F_BPAE_html=F_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_BPAE_html=T_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_BPAE_html=D_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
F_PACE_html=F_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_PACE_html=T_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_PACE_html=D_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')

FBI1_html=FBI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
FBI2_html=FBI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TBI1_html=TBI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TBI2_html=TBI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DBI1_html=DBI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DBI2_html=DBI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')

FPI1_html=FPI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
FPI2_html=FPI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TPI1_html=TPI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TPI2_html=TPI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DPI1_html=DPI1_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DPI2_html=DPI2_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')


# row color & padding
F_BPAE_html=F_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_BPAE_html=T_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_BPAE_html=D_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
F_PACE_html=F_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_PACE_html=T_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_PACE_html=D_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')

FBI1_html=FBI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
FBI2_html=FBI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TBI1_html=TBI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TBI2_html=TBI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DBI1_html=DBI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DBI2_html=DBI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')

FPI1_html=FPI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
FPI2_html=FPI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TPI1_html=TPI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TPI2_html=TPI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DPI1_html=DPI1_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DPI2_html=DPI2_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')


# model, tool
F_BPAE_html=F_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
F_BPAE_html=F_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
T_BPAE_html=T_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
T_BPAE_html=T_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
D_BPAE_html=D_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
D_BPAE_html=D_BPAE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')

F_PACE_html=F_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
F_PACE_html=F_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
T_PACE_html=T_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
T_PACE_html=T_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')
D_PACE_html=D_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Tool</th>')
D_PACE_html=D_PACE_html.replace('<th>Tool</th>','<th style="background-color: aqua;">Model</th>')


#this week remark
thisweek_html='<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">'+this_week+'</th>'
thisweek_strong='<th style="text-align: center; background-color:rgb(192, 0, 0); color:white; padding:5px;">'+this_week+'</th>'
F_BPAE_html=F_BPAE_html.replace(thisweek_html,thisweek_strong)
T_BPAE_html=T_BPAE_html.replace(thisweek_html,thisweek_strong)
D_BPAE_html=D_BPAE_html.replace(thisweek_html,thisweek_strong)
F_PACE_html=F_PACE_html.replace(thisweek_html,thisweek_strong)
T_PACE_html=T_PACE_html.replace(thisweek_html,thisweek_strong)
D_PACE_html=D_PACE_html.replace(thisweek_html,thisweek_strong)
