import pandas as pd
from flask import Flask, render_template, flash, url_for, request, make_response, jsonify, session,send_from_directory
from werkzeug.utils import secure_filename
import os, time
import io
import base64
import json
import datetime
import xlsxwriter
import os, sys, glob
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.styles import Border, Side
import openpyxl
from openpyxl.styles import Border, Side, Font, Alignment
from openpyxl.styles import Border, Side
#from flask import send_from_directory
from flask import send_file
import smtplib,ssl
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from fnmatch import fnmatch
import time
from Report_Genration import Report_Genration

working_directory ="AIR_VELOCITY_REPORT\\{}"
final_working_directory ="AIR_VELOCITY_REPORT\\{}\\{}.xlsx"

app = Flask(__name__)
app.secret_key = 'file_upload_key'

MYDIR = os.path.dirname(__file__)
app.config['UPLOAD_FOLDER'] = "static/inputData/"

equipment_master  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"equipment_master.xlsx"))
ISO_guidlines_master  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"ISO_guidlines.xlsx"))
EUGMP_guidlines_master  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"EUGMP_guidlines.xlsx"))
equipment_master['SR_NO_ID']  = equipment_master['SR_NO_ID'].astype(str)
equipment_master['DUE_DATE']  = equipment_master['DUE_DATE'].astype(str)
equipment_master['DONE_DATE'] = equipment_master['DONE_DATE'].astype(str)




guidlance_list                = ISO_guidlines_master.Guidelines.unique().tolist()+EUGMP_guidlines_master.Guidelines.unique().tolist()
serial_id_list_pao            = equipment_master[equipment_master['Type']=='PAO_TEST'].SR_NO_ID.unique().tolist()
serial_id_list_particle_count = equipment_master[equipment_master['Type']=='PARTICLE_COUNT'].SR_NO_ID.unique().tolist()
serial_id_list_air_velocity   = equipment_master[equipment_master['Type']=='AIR_VELOCITY'].SR_NO_ID.unique().tolist()

sent_mail                     = False

condition_list                = ['At Rest','In Operation']
grade_list                    = ['A','B','C','D']


server    = 'smtp.gmail.com'
port      =  587
username  =  "aajeetshk@gmail.com"
password  =  "ilbumnmnsnqletdk"
send_from = "aajeetshk@gmail.com"
send_to   = "ashish@pinpointengineers.co.in"

def send_mail(subject,text,files,file_name,isTls=True):
        msg = MIMEMultipart()
        msg['From'] = send_from
        msg['To'] = send_to
        msg['Date'] = formatdate(localtime = True)
        msg['Subject'] = subject
        msg.attach(MIMEText(text))

        part = MIMEBase('application', "octet-stream")
        part.set_payload(open(files, "rb").read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', 'attachment; filename={}.xlsx'.format(file_name))
        msg.attach(part)
        
        #context = ssl.SSLContext(ssl.PROTOCOL_SSLv3)
        #SSL connection only working on Python 3+
        smtp = smtplib.SMTP(server, port)
        if isTls:
            smtp.starttls()
        smtp.login(username,password)
        smtp.sendmail(send_from, send_to, msg.as_string())
        smtp.quit()
@app.route("/")
def default():
    return make_response(render_template('login_page/login.html'),200)    
    
@app.route("/login", methods=["GET", "POST"])
def login():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    company_name_list             = customer_details.COMPANY_NAME.unique().tolist()
    if request.method == 'POST':
      form_data = request.form
      l_id = form_data['login']
      pwd = form_data['password']
      if(l_id.lower()=='admin'.lower() and pwd == 'admin'):
          print('inside if')
          session['username'] = l_id
          flash('Login Successful')
          return make_response(render_template('Air_velocity.html',company_list=company_name_list,
                                grade_list=grade_list,
                                equipment_list =serial_id_list_air_velocity,
                                msg = True, err = False, warn = False),200)
      else:
          print('inside else')
          flash('Invalid Credentials')
          return make_response(render_template("login_page/login.html", msg = False, err = True, warn = False),403)
    else:
        print('get request')
        
@app.route('/logout')
def logout():
    session.pop('username', None)
    session.pop('selected_file', None)
    global Selected_files
    Selected_file = None
    flash('Logout Successful')
    return make_response(render_template("login_page/login.html",msg = True, err = False, warn = False, message='Logout Successful'),200)

@app.route("/Air_velocity")
def Air_velocity():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    company_name_list             = customer_details.COMPANY_NAME.unique().tolist()
    return make_response(render_template('Air_velocity.html',grade_list=grade_list,
    company_list=company_name_list,equipment_list =serial_id_list_air_velocity),200) 
    
@app.route("/paotest")
def paotest():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    company_name_list             = customer_details.COMPANY_NAME.unique().tolist()
    return make_response(render_template('PAO.html',company_list=company_name_list,
                            equipment_list =serial_id_list_pao),200)
                            
@app.route("/UpdateCompanyDetails")
def UpdateCompanyDetails():   
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    customer_list   = customer_details.to_dict('records')
    return make_response(render_template('UpdateCompanyDetails.html',customer_list  = customer_list),200) 
    
@app.route("/consolidation")
def consolidation():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    company_name_list             = customer_details.COMPANY_NAME.unique().tolist()
    return make_response(render_template('consolidation.html'),200)
    
@app.route("/particle_count")
def particle_count():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    company_name_list             = customer_details.COMPANY_NAME.unique().tolist()
    return make_response(render_template('particle_count.html',company_list=company_name_list,
					 guidlance_list=guidlance_list  ,
					  equipment_list =serial_id_list_particle_count,
                                          condition_list  =condition_list ),200)
 
@app.route("/get_available_directory", methods=['POST', 'GET'])
def get_available_directory():
    data          = request.args.get('params_data')
    report_type   = json.loads(data)
    report_type   = report_type.replace(" ","_")
    directory     = 'static\\Report\\{}'.format(report_type)
    sub_list      =os.listdir(directory)
    sub_list.insert(0,"ALL")
    dict_list=[]
    for x in sub_list:
        thisdict ={"id":x,"name":x}
        dict_list.append(thisdict)
    d = {"dict_list":dict_list,
         "start_date":str(datetime.datetime.today().strftime('%d/%m/%Y')),
         "end_date":str(datetime.datetime.today().strftime('%d/%m/%Y')),        
        }   
    return json.dumps(d)    
    
 
@app.route("/update_company_details", methods=['POST', 'GET'])
def update_company_details():
    customer_details  = pd.read_excel(os.path.join(app.config['UPLOAD_FOLDER'],"company_details.xlsx"))
    data          = request.args.get('params_data')
    company_name_val     = json.loads(data)
    company_address = ""
    report_id       = ""    
    location        = ""
    test_taken      = datetime.datetime.today().strftime('%d/%m/%Y')
    if company_name_val is not None :
        temp_df = customer_details.loc[(customer_details.COMPANY_NAME==company_name_val)]
        report_id =str(temp_df.REPORT_NUMBER.values[0])
        company_address =str(temp_df.ADDRESS.values[0])       
       
    d = {
        "error":"none",
        "company_address":company_address,
        "report_id":str("PPE0{}AV01A".format(report_id)),
        "test_taken":test_taken,
        "location":location,
    }
    return json.dumps(d) 
    
@app.route("/update_instument_details", methods=['POST', 'GET'])
def update_instument_details():
    
    data          = request.args.get('params_data')
    SR_NO_val     = json.loads(data)
    print(SR_NO_val)
    INSTRUMENT_NAME = ""
    MAKE           = ""
    MODEL_NUMBER   = "" 
    done_date      = ""
    due_date       = ""
    VALIDITY        = ""

    if SR_NO_val is not None :
        temp_df = equipment_master.loc[(equipment_master.SR_NO_ID==SR_NO_val)]       
        INSTRUMENT_NAME = str(temp_df.EQUIPMENT_NAME.values[0])
        MAKE           = str(temp_df.MAKE.values[0])
        MODEL_NUMBER   = str(temp_df.MODEL_NUMBER.values[0])
        done_date      = str(temp_df.DONE_DATE.values[0])
        due_date       = str(temp_df.DUE_DATE.values[0])
        VALIDITY        = str(temp_df.DUE_DATE.values[0]).replace("-","/")
        
       
    d = {
        "error":"none",
        "INSTRUMENT_NAME" : INSTRUMENT_NAME,
        "MAKE"           :  MAKE,
        "MODEL_NUMBER"   :  MODEL_NUMBER,
        "done_date"      :  done_date,
        "due_date"       :  due_date,
        "VALIDITY"        : VALIDITY,
        

    }
    return json.dumps(d)  



@app.route("/update_grade", methods=['POST', 'GET'])
def update_grade():    
    data          = request.args.get('params_data')
    gl_value      = json.loads(data)  
    print(gl_value)
    if "ISO" in gl_value :    
        grade_list    = ISO_guidlines_master.loc[(ISO_guidlines_master.Guidelines==gl_value)]['Grade'].tolist()      
    if "EU" in gl_value  :
        grade_list    = EUGMP_guidlines_master.loc[(EUGMP_guidlines_master.Guidelines==gl_value)]['Grade'].unique().tolist()      
      
    dict_list=[]
    for x in grade_list:
        thisdict ={"id":x,"name":x}
        dict_list.append(thisdict)
    d = {"dict_list":dict_list,"error":"none"}
    return json.dumps(d)      

@app.route("/get_limits", methods=['POST', 'GET'])
def get_limits():    
    data          = request.args.get('params_data')
    full_data     = json.loads(data)   
    gl_value      = full_data['gl_value']
    grade         = full_data['grade']
    condition     = full_data['condition']
    print(gl_value)
    if "EU" not in gl_value :
        value1        = ISO_guidlines_master.loc[(ISO_guidlines_master.Guidelines==gl_value)
                                                    &
                                            (ISO_guidlines_master.Grade==grade)
                                             ]['point_five_percent'].values[0]
        value2        = ISO_guidlines_master.loc[(ISO_guidlines_master.Guidelines==gl_value)
                                                    &
                                            (ISO_guidlines_master.Grade==grade)
                                             ]['five_percent'].values[0]   
    if "EU" in gl_value  : 
    
        value1        = EUGMP_guidlines_master.loc[(EUGMP_guidlines_master.Guidelines==gl_value)
                                                    &
                                            (EUGMP_guidlines_master.Grade==grade)
                                                    &
                                            (EUGMP_guidlines_master.Condition==condition)
                                             ]['point_five_percent'].values[0]
                                             
                                             
        value2        = EUGMP_guidlines_master.loc[(EUGMP_guidlines_master.Guidelines==gl_value)
                                                    &
                                            (EUGMP_guidlines_master.Grade==grade)
                                                    &
                                            (EUGMP_guidlines_master.Condition==condition)
                                             ]['point_five_percent'].values[0]
    
        
        
    
    d = {"value1":str(value1),
         "value2":str(value2),
         "error":"none"}
    print(d)
    return json.dumps(d)      


@app.route("/submit_data")
def submit_data():
   
    data          = request.args.get('params_data')
    full_data     = json.loads(data)
    basic_details = full_data['basic_details']
    observation   = full_data['observation']    
    company_name  = basic_details['company_name']
    temp_df       = pd.DataFrame.from_dict(observation,orient ='index')
    file_name,file_path=Report_Genration.generate_report(temp_df ,basic_details)    
    subject       = "HVAC-Air Velocity Automated Genrated Report - {}".format(company_name)
    text          = "Hi PinPoint Team \n\nPlease find attached automated Generated File {} for {} \n\nRegards \nAjeet Shukla :) :) :)".format(file_name,company_name)
   
    if sent_mail:
        send_mail(subject,text,file_path,file_name) 
    d = {
        "error":"none",
        "file_name":file_name,
        "file_path":file_path
        }
   
    return json.dumps(d)
    
    
@app.route("/submit_consolidated")
def submit_consolidation():


    data           = request.args.get('params_data')
    full_data      = json.loads(data)
    start_date     = full_data['start_date']
    end_date       = full_data['end_date']
    report_type    = full_data['report_type']
    company_name   = full_data['company_name']
    report_type    = report_type.replace(" ","_")
    directory      = 'static\\{}'.format(report_type)
    
    mypath     = 'static\\Report\\{}\{}'.format(report_type,company_name)
    print(company_name)
    if company_name=="ALL":
        mypath = 'static\\Report\\{}'.format(report_type)
    
    pattern = "*.xlsx"


    print(start_date)
    print(end_date)
    datetime_object_start = datetime.datetime.strptime(start_date, '%Y-%m-%d')
    datetime_object_end   = datetime.datetime.strptime(end_date, '%Y-%m-%d')

    file_list= []
    for path, subdirs, files in os.walk(mypath):
        for name in files:
            if fnmatch(name, pattern):          
                file_time = time.strftime('%d/%m/%Y',time.gmtime(os.path.getmtime(os.path.join(path, name) )))
                datetime_object = datetime.datetime.strptime(file_time, '%d/%m/%Y')
                if datetime_object_start<=datetime_object<=datetime_object_end :
                    file_list.append(os.path.join(path, name))
    print(file_list)
    
    d={"msg":"done",}
    return json.dumps(d)
    
    
@app.route("/submit_data_pao")
def submit_data_pao():
    print("submitdata called pao")
    data          = request.args.get('params_data')
    full_data     = json.loads(data)
    basic_details = full_data['basic_details']
    observation   = full_data['observation']
    company_name  = basic_details['company_name']
    temp_df = pd.DataFrame.from_dict(observation,orient ='index')
    file_name,file_path=Report_Genration.generate_report_pao(temp_df,basic_details)
    subject   = "HVAC-PAO Automated Genrated Report - {}".format(company_name)
    text      = "Hi PinPoint Team \n\nPlease find attached automated Generated File {} for {} \n\nRegards \nAjeet Shukla :) :) :)".format(file_name,company_name)

    if sent_mail :
        send_mail(subject,text,file_path,file_name) 
    d = {
        "error":"none",
        "file_name":file_name,
        "file_path":file_path
        }
    
    return json.dumps(d)
 
@app.route("/submit_particle_report")
def submit_particle_report():
    data          = request.args.get('params_data')
    full_data     = json.loads(data)
    basic_details = full_data['basic_details']
    observation   = full_data['observation']	
    company_name  = basic_details['company_name']
    temp_df       = pd.DataFrame.from_dict(observation,orient ='index')
    file_name,file_path=Report_Genration.generate_report_particle_count(temp_df,basic_details)
    subject   = "Particle Count Automated Genrated Report - {}".format(company_name)
    text      = "Hi PinPoint Team \n\nPlease find attached automated Generated File {} for {} \n\nRegards \nAjeet Shukla :) :) :)".format(file_name,company_name)
    if sent_mail:
        send_mail(subject,text,file_path,file_name) 
    d = {
        "error":"none",
        "file_name":file_name,
        "file_path":file_path
        }
    
    return json.dumps(d)
    
    
@app.route("/submit_updateCompanyDetails")    
def submit_updateCompanyDetails():   
    data            = request.args.get('params_data')
    data            = json.loads(data)  
    observation     = data['observation']
    temp_df         = pd.DataFrame.from_dict(observation,orient ='index')
    temp_df =temp_df[['COMPANY_NAME','ADDRESS','REPORT_NUMBER']]
 
    final_working_directory = MYDIR + "static/inputData/company_details.xlsx"
    temp_df.to_excel(final_working_directory,index=False)
    d = {
        "error":"none",
        }
    
    return json.dumps(d)
   

if __name__ == '__main__':
    app.debug = True
    app.run()

