from Adata import F_BPAE_Merge,T_BPAE_Merge,D_BPAE_Merge,F_PACE_Merge,T_PACE_Merge,D_PACE_Merge,FBI_merge,TBI_merge,DBI_merge,FPI_merge,TPI_merge,DPI_merge, this_week

############################ html  ############################  
# change to html -> table & border
F_BPAE_html=F_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
T_BPAE_html=T_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
D_BPAE_html=D_BPAE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
F_PACE_html=F_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
T_PACE_html=T_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
D_PACE_html=D_PACE_Merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:1rem;border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
FBI_html=FBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TBI_html=TBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DBI_html=DBI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
FPI_html=FPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
TPI_html=TPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')
DPI_html=DPI_merge.to_html().replace('<table border="1"','<table border="1" style="border:1px solid rgb(188, 188, 188); font-size:0.8rem; border-collapse:collapse; text-align:center;font-family:Arial Narrow;"')


# text align center & column color
F_BPAE_html=F_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_BPAE_html=T_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_BPAE_html=D_BPAE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
F_PACE_html=F_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
T_PACE_html=T_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')
D_PACE_html=D_PACE_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(238, 238, 238);">')

FBI_html=FBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TBI_html=TBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DBI_html=DBI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
FPI_html=FPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
TPI_html=TPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')
DPI_html=DPI_html.replace('<tr style="text-align: right;">','<tr style="text-align: center; background-color:rgb(221,235,247);">')


# row color & padding
F_BPAE_html=F_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_BPAE_html=T_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_BPAE_html=D_BPAE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
F_PACE_html=F_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
T_PACE_html=T_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
D_PACE_html=D_PACE_html.replace('<th>','<th style="text-align: center; background-color:rgb(238, 238, 238); padding:5px;">')
FBI_html=FBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TBI_html=TBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DBI_html=DBI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
FPI_html=FPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
TPI_html=TPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')
DPI_html=DPI_html.replace('<th>','<th style="text-align: center; background-color:rgb(221,235,247); padding:5px;">')


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
