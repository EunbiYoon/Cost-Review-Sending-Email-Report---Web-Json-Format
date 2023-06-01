import pandas as pd
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
from Chtml import this_week, F_BPAE_html, T_BPAE_html, D_BPAE_html, F_PACE_html, T_PACE_html, D_PACE_html,FBI1_html,FBI2_html,TBI1_html,TBI2_html,DBI1_html,DBI2_html,FPI1_html,FPI2_html,TPI1_html,TPI2_html,DPI1_html,DPI2_html
from Bmatgrptbl import save_path,today_date
import os
import Bmatgrptbl

#html - table
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()


#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
# msg['To' ]='iggeun.kwon@lge.com'
# msg['Cc']='janine.williams@lge.com, karina1.beveridge@lge.com, kitae3.park@lge.com, soyoung1.an@lge.com, soyoon1.kim@lge.com, wolyong.ha@lge.com, grace.hwang@lge.com, tg.kim@lge.com, seongju.yu@lge.com, minhyoung.sun@lge.com, jongseop.kim@lge.com, richard.song@lge.com, gilnam.lee@lge.com, jacey.jung@lge.com, ##312718@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#Subject 꾸미기
msg['Subject']='Cost Review Report '+this_week
msg.attach(MIMEText('<h4 style="font-family:sans-serif; font-weight:500">Dear all,<br/><br/>I would like to share this week cost review report and detailed informations are in below website.<br/>You can access website from CloudPC or LG wifi for security purpose. <a href="http://10.225.2.85">http://10.225.2.85</a></h4>','html'))
# msg.attach(MIMEText('<h3 style="font-family:sans-serif;">[matching number format in item list]<br/>Dear Catalina,</h3><h4 style="font-family:sans-serif; font-weight:500">Please chek final data of item and graph change,</h4>','html'))


# graph file read
with open(save_path+"FL_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image1 = MIMEImage(img_data, name=os.path.basename("FL_BPA_Entity"+today_date+'.png'))

with open(save_path+"TL_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image2 = MIMEImage(img_data, name=os.path.basename("TL_BPA_Entity"+today_date+'.png'))

with open(save_path+"DR_BPA_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image3 = MIMEImage(img_data, name=os.path.basename("DR_BPA_Entity"+today_date+'.png'))


with open(save_path+"FL_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image4 = MIMEImage(img_data, name=os.path.basename("FL_PAC_Entity"+today_date+'.png'))

with open(save_path+"TL_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image5 = MIMEImage(img_data, name=os.path.basename("TL_PAC_Entity"+today_date+'.png'))

with open(save_path+"DR_PAC_Entity"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image6 = MIMEImage(img_data, name=os.path.basename("DR_PAC_Entity"+today_date+'.png'))


# html table attach
F_BPAE_attach = MIMEText(F_BPAE_html, "html")
T_BPAE_attach = MIMEText(T_BPAE_html, "html")
D_BPAE_attach = MIMEText(D_BPAE_html, "html")

F_PACE_attach = MIMEText(F_PACE_html, "html")
T_PACE_attach = MIMEText(T_PACE_html, "html")
D_PACE_attach = MIMEText(D_PACE_html, "html")

FBI1_attach = MIMEText(FBI1_html, "html")
FBI2_attach = MIMEText(FBI2_html, "html")
TBI1_attach = MIMEText(TBI1_html, "html")
TBI2_attach = MIMEText(TBI2_html, "html")
DBI1_attach = MIMEText(DBI1_html, "html")
DBI2_attach = MIMEText(DBI2_html, "html")

FPI1_attach = MIMEText(FPI1_html, "html")
FPI2_attach = MIMEText(FPI2_html, "html")
TPI1_attach = MIMEText(TPI1_html, "html")
TPI2_attach = MIMEText(TPI2_html, "html")
DPI1_attach = MIMEText(DPI1_html, "html")
DPI2_attach = MIMEText(DPI2_html, "html")


msg.attach(MIMEText('<br/><h3 style="font-family:sans-serif;">Front Loader BPA Entity Trend</h3>','html'))
msg.attach(image1)
msg.attach(F_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FBI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FBI2_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Top Loader BPA Entity Trend</h3>','html'))
msg.attach(image2)
msg.attach(T_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TBI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TBI2_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Dryer BPA Entity Trend</h3>','html'))
msg.attach(image3)
msg.attach(D_BPAE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DBI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DBI2_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Front Loader PAC Entity Trend</h3>','html'))
msg.attach(image4)
msg.attach(F_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FPI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(FPI2_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Top Loader PAC Entity Trend</h3>','html'))
msg.attach(image5)
msg.attach(T_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TPI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(TPI2_attach)

msg.attach(MIMEText('<br/><br/><h3 style="font-family:sans-serif;">Dryer PAC Entity Trend</h3>','html'))
msg.attach(image6)
msg.attach(D_PACE_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DPI1_attach)
msg.attach(MIMEText('<h6> </h6>','html'))
msg.attach(DPI2_attach)


#첨부 파일1
etcFileName='Cost Review_0602.xlsx'
with open("C:/Users/RnD Workstation/Documents/CostReview/0602/Cost Review_0602.xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")