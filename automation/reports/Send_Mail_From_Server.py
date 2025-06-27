import zipfile
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase 
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email import encoders
import smtplib
from email.mime.text import MIMEText
import os
import shutil

#email_to and email_cc son listas
def send_mail_img_files(subject,email_to,email_cc,img_path,files_path,zip_path):
    #try:
    emails_to_send = []

    for i in email_to:
        emails_to_send.append(i)
    for i in email_cc:
        emails_to_send.append(i)

    
    name = files_path
    zip_name = zip_path + '\\Tool_Results.zip'

    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
        for folder_name, subfolders, filenames in os.walk(name):
            for filename in filenames:
                file_path = os.path.join(folder_name, filename)
                zip_ref.write(file_path, arcname=os.path.relpath(file_path, name))

    with open(img_path, 'rb') as f:
        img_data = f.read()

    msg = MIMEMultipart()
    attachment = zip_name

    msg['From'] = "correo@gmail.com"
    msg['To'] = ','.join(email_to)
    if len(email_cc) > 0:
        msg['Cc'] = ','.join(email_cc)
    msg['Subject'] = subject

    #html = f"""<p style="color:Black;font-size:16px;background-color: rgb(240, 255, 12);"><i><b>{text_email}</b></i></p>""".format(text_email = text_email)

    part = MIMEBase('application', "octet-stream")
    part.set_payload(open(attachment, "rb").read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment', filename="Tool_Results.zip")

    img = MIMEImage(img_data)
    img.add_header('Content-ID', '<image1>')
    msg.attach(img)

    # Agregar cuerpo del correo
    html = '<img src="cid:image1" class="responsive">'
    #html += '<br><img src="cid:image1" class="responsive"><br>'

    msg.attach(MIMEText(html, 'html'))
    msg.attach(part)

    server = smtplib.SMTP("smtp.huawei.com", 587)
    #Next, log in to the server
    server.starttls()
    server.login(r'user','password')
    text = msg.as_string()
    server.sendmail("correo@gmail.com",emails_to_send, text)
    print("Email sent")
    server.quit()

    try:
        os.remove(zip_name)
    except:
        print('no zip results file')
    #except:
    #    print('Error al mandar email')

def send_mail_files(subject,email_to,email_cc,files_path,zip_path,text_email):
    try:
        emails_to_send = []

        for i in email_to:
            emails_to_send.append(i)
        for i in email_cc:
            emails_to_send.append(i)

        
        name = files_path
        zip_name = zip_path + '\\results.zip'

        with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
            for folder_name, subfolders, filenames in os.walk(name):
                for filename in filenames:
                    file_path = os.path.join(folder_name, filename)
                    zip_ref.write(file_path, arcname=os.path.relpath(file_path, name))

        msg = MIMEMultipart()
        attachment = zip_name

        msg['From'] = "correo@gmail.com"
        msg['To'] = ','.join(email_to)
        if len(email_cc) > 0:
            msg['Cc'] = ','.join(email_cc)
        msg['Subject'] = subject

        html = f"""<p style="color:Black;font-size:16px;background-color: rgb(240, 255, 12);"><i><b>{text_email}</b></i></p>""".format(text_email = text_email)

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(attachment, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment', filename="results.zip")


        # Agregar cuerpo del correo
    

        msg.attach(MIMEText(html, 'html'))
        msg.attach(part)

        server = smtplib.SMTP("smtp.huawei.com", 587)
        #Next, log in to the server
        server.starttls()
        server.login(r'user','password')
        text = msg.as_string()
        server.sendmail("correo@gmail.com",emails_to_send, text)
        print("Email sent")
        server.quit()

        try:
            os.remove(zip_name)
        except:
            print('no zip results file')
    except:
        print('Error al mandar email')

def send_mail_text(subject,email_to,email_cc,text_email):
    try:
        emails_to_send = []

        for i in email_to:
            emails_to_send.append(i)
        for i in email_cc:
            emails_to_send.append(i)

        msg = MIMEMultipart()

        msg['From'] = "correo@gmail.com"
        msg['To'] = ','.join(email_to)
        if len(email_cc) > 0:
            msg['Cc'] = ','.join(email_cc)
        msg['Subject'] = subject

        html = f"""<p style="color:Black;font-size:16px;background-color: rgb(240, 255, 12);"><i><b>{text_email}</b></i></p>""".format(text_email = text_email)

        # Agregar cuerpo del correo 
        msg.attach(MIMEText(html, 'html'))

        server = smtplib.SMTP("smtp.huawei.com", 587)
        #Next, log in to the server
        server.starttls()
        server.login(r'user','password')
        text = msg.as_string()
        server.sendmail("correo@gmail.com",emails_to_send, text)
        print("Email sent")
        server.quit()
    except:
        print('Error al mandar email')
