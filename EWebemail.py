import pandas as pd
import smtplib
import email
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication 
# from Adata import this_week
import os

# today date
today_date='0714'
this_week='23.07 W2'

#html - table
server = smtplib.SMTP('lgekrhqmh01.lge.com:25')
server.ehlo()

#메일 내용 구성
msg=MIMEMultipart()

# 수신자 발신자 지정
msg['From']='eunbi1.yoon@lge.com'
msg['To' ]='daewook.kwak@lge.com'
msg['Cc']='janine.williams@lge.com, karina1.beveridge@lge.com, kitae3.park@lge.com, soyoung1.an@lge.com, soyoon1.kim@lge.com, wolyong.ha@lge.com, grace.hwang@lge.com, tg.kim@lge.com, seongju.yu@lge.com, minhyoung.sun@lge.com, jongseop.kim@lge.com, richard.song@lge.com, gilnam.lee@lge.com, jacey.jung@lge.com, ##312718@lge.com'
msg['Bcc']='eunbi1.yoon@lge.com'

#Subject 꾸미기
msg['Subject']='Cost Review Report '+this_week
msg.attach(MIMEText('<h4 style="font-family:sans-serif; font-weight:500">Dear all,<br/><br/>I would like to share this week cost review report and detailed informations are in below website.<br/>You can access website with Chrome or Edge in CloudPC or LG wifi for security purpose. <a href="http://10.225.2.85">http://10.225.2.85</a></h4>','html'))

save_path='C:/Users/RnD Workstation/Documents/CostReview/'+today_date+'/'

# graph file read
with open(save_path+"FL_BPA_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image1 = MIMEImage(img_data, name=os.path.basename("FL_BPA_Entity"+today_date+'.jpg'))

with open(save_path+"TL_BPA_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image2 = MIMEImage(img_data, name=os.path.basename("TL_BPA_Entity"+today_date+'.jpg'))

with open(save_path+"DR_BPA_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image3 = MIMEImage(img_data, name=os.path.basename("DR_BPA_Entity"+today_date+'.jpg'))


with open(save_path+"FL_PAC_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image4 = MIMEImage(img_data, name=os.path.basename("FL_PAC_Entity"+today_date+'.jpg'))

with open(save_path+"TL_PAC_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image5 = MIMEImage(img_data, name=os.path.basename("TL_PAC_Entity"+today_date+'.jpg'))

with open(save_path+"DR_PAC_Entity_"+today_date+'.png', 'rb') as f:
        img_data = f.read()
image6 = MIMEImage(img_data, name=os.path.basename("DR_PAC_Entity"+today_date+'.jpg'))





# msg.attach(MIMEText('<br/><h3 style="font-family:sans-serif;">Front Loader BPA Entity Trend</h3>','html'))
msg.attach(image1)
msg.attach(image2)
msg.attach(image3)
msg.attach(image4)
msg.attach(image5)
msg.attach(image6)



#첨부 파일1
etcFileName='Cost Review_'+today_date+'.xlsx'
with open("C:/Users/RnD Workstation/Documents/CostReview/"+today_date+"/Cost Review_"+today_date+".xlsx", 'rb') as etcFD : 
    etcPart = MIMEApplication( etcFD.read() )
    #첨부파일의 정보를 헤더로 추가
    etcPart.add_header('Content-Disposition','attachment', filename=etcFileName)
    msg.attach(etcPart)


#메세지 보내고 확인하기
server.send_message(msg)
server.close()
print("Sucess!!!")