
File_Name = 'Ontodocs Controls'
sheet = "Uptime Monitor"
Column_No = 3
Column = "C"

# check column length
def Column_Length_Silent(File_Name,Column,Column_No,sheet):
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',\
                    "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json',scope)
        client = gspread.authorize(creds)
        sheet = client.open(File_Name).worksheet(sheet)
        lists = sheet.col_values(Column_No)
    
### Consider what is happening on the e-mail column
        Len_all = len(sheet.get_all_values())
        cell_list = sheet.range(str(Column)+'1:'+str(Column)+str(Len_all)+'')
        cell_list_re = sheet.range(str(Column)+'1:'+str(Column)+str(Len_all)+'')
#
        length = len(cell_list) 
        i = 0
   # 
   
    # Iterating using while loop 
        while i < length: 
            cell_list[i] = str(cell_list[i]).replace('>', '').split()[2]
        #print(cell_list) 
            i += 1
        return(cell_list,sheet)


# Script to update cell with specific words

def sheet_updater_silent(File_Name,sheet1,Column,Column_No):
    
    ################# Inumerate to confirm if we can send out e-mail #######################

    
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    sheet1 = sheet
    cell_list = Column_Length_Silent(File_Name,Column,Column_No,sheet) # the columns of the "sent" reference
    sheet4 = cell_list[1]
    sent_no = len([i for i, s in enumerate(cell_list[0]) if 'Timed' in s]) 
    tosend_no = len([i for i, s in enumerate(cell_list[0]) if 'Timed' not in s])
    
    if tosend_no !=0:
        sent = [i for i, s in enumerate(cell_list[0]) if 'Timed' in s]
        tosend = [i for i, s in enumerate(cell_list[0]) if 'Timed' not in s]
        row_num = (tosend[0]+1 )#+ nottosend[0])
        #row_num = (tosend[0] )#+ nottosend[0])
    
    
    #for val in tosend:
     #   if val == 0:
      #      break
        
        

    
    #print("Can proceed with generating output:")
    #Market_Report(row_num)
        #Dupont_Doc_Row(data_row) # - this is crux of ontodocs - the document creation function **CRUX-ONTODOCS**
     #   Yes_sendmail(Co_name,to_email) # still dont understand where it gets the e-mail from
      #  sheet4.update_cell(row_num, 16, 'sent A')
    
    #sheet4.update_cell(row_num, 4, 'sent Z')



    #print("Cell_List :",cell_list[0],len(cell_list[0]))
    #print("Sent No :",sent_no) # items where there is a "sent"
    #print("To Send No :",tosend_no) # items where there is a "sent"
    #print("To Send No :",tosend_no) # items where there is a "sent"
    #print("Next Row Number to update",row_num)
    #data_row = sheet4.row_values(row_num)
    #Co_name = data_row[0] #- correct
    #to_email = data_row[1]
    #products_no = data_row[3] # cant be blank
    #print("The e-mail goes to:",to_email," for product(s) quantity:",products_no)
    
    
    
    ################# Function Pull the product for pdf, before sending e-mail #######################

    
    #NoProducts = 2
    #Pull_LatestQuote2(products_no)

    
    ################# Function Pull the product for pdf, before sending e-mail #######################

    #Co_name = 'Bofin Consulting'
    #Yes_sendmail(Co_name,to_email)
    
    
    ################# Update the row that we have sent out the e-mail #######################

    now = datetime.now() 
    run_time = datetime.now() + timedelta(hours=2)
    dt_string = run_time.strftime("%d/%m/%Y %H:%M:%S")
    #print("No items to sent: error or no emails to send:", dt_string)	

    sheet4.update_cell(row_num, (ord(Column)-64), 'Timed:'+str(dt_string)) #16 is equivalent to column"P"
    #sheet4.update_cell(row_num,  'Timed at 12:00') #16 is equivalent to column"P"
    
    #print("Succeessful: We have updated the uptime")



#!pip install gspread

# Load the appropriate libraries

import gspread
from oauth2client.service_account import ServiceAccountCredentials
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',\
         "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json',scope)
client = gspread.authorize(creds)



#!pip install dropbox
import dropbox
dbx = dropbox.Dropbox('WjlGCYBkvxcAAAAAAAAAAZ0y6JqSSylGLQDXj1hVI-kSExF7yAmuYomoe9qvsGyn')  #- ChatsOntodocs


#!pip install convertapi
import convertapi
convertapi.api_secret = 'lt0g2TvEr0MRrWJ2'

import time

from datetime import date
from datetime import datetime
from datetime import timedelta  


import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

#!pip install schedule

#!pip install schedule
import schedule

### Time and scheduler modules
import schedule
import time

from datetime import date
from datetime import datetime
from datetime import timedelta 



def Yes_mailbody1(OTP): # This function determines teh e-mail structure
    body = '<p>Good Day</p>\
        <p>Please find attached the Quotation you requested.</p>\
        <p>The Ouote Reference is :  '+'<b>'+OTP+'</b>'+ ' for future use.</p>\
        <p>Please do not hesitate to contact us for further information.</p>\
        <p>Kind Regards</p>\
        <p><strong>The <span style="color: #ff9900;">Outsource Practice</span></strong></p>'
    return(body)


def Yes_OTPsendmail(OTP,to_email):

    fromaddr = 'info@outsourcepractice.co.za'
    toaddr = to_email

    msg = MIMEMultipart()

    msg['From'] = fromaddr
    msg['To'] = toaddr
    msg['Subject'] = 'Quotation Request '

    body = Yes_mailbody1(OTP)
        
    msg.attach(MIMEText(body, 'html'))

    filename = 'Quotation.pdf' ### test name - pdf1
    attachment = open(str('../Chatbooks/Quotation2.pdf'), "rb") ### Remmber this test name - pdf1

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename= "%s"' % filename)

    msg.attach(part)

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(fromaddr, 'Simple101*')
    text = msg.as_string()
    server.sendmail(fromaddr, toaddr, text)
    server.quit()
    

 #Obtain the column length for teh ontodocts sheet 
def Column_Length_OTP(Column,sheet):
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',\
                    "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json',scope)
        client = gspread.authorize(creds)
        sheet4 = client.open("Chatbooks QB Manager").worksheet(sheet)
        lists = sheet4.col_values(21)
    
### Consider what is happening on the e-mail column
        Len_all = len(sheet4.get_all_values())
        cell_list = sheet4.range(str(Column)+'1:'+str(Column)+str(Len_all)+'')
        cell_list_re = sheet4.range(str(Column)+'1:'+str(Column)+str(Len_all)+'')
#
        length = len(cell_list) 
        i = 0
   # 
   
    # Iterating using while loop 
        while i < length: 
            cell_list[i] = str(cell_list[i]).replace('>', '').split()[2]
        #print(cell_list) 
            i += 1
        return(cell_list,sheet4)
    
def Pull_LatestQuote_clean(NoProducts):

    ## Cross reference the report name to the last doc in the excel file
    
    ## Connect to the google shreadsheet
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',\
             "https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name('creds.json',scope)
    client = gspread.authorize(creds)

# Idenify the last document [select_doc] to convert and e-mail

    sheet3 = client.open("Chatbooks QB Manager").worksheet('Quotations Manager')
    doc_ids = sheet3.col_values(3)
    email_ids = sheet3.col_values(2)

    select_doc = doc_ids[len(doc_ids)-1] # this number determines second last or last (-1 = last, -2 = second last)
    select_email = email_ids[len(email_ids)-1]

    #print(select_doc) # this is the last reference on row
    #print(select_email) 

    reportA = '/'+str(select_doc)+'.xlsx'
    report = str(select_doc)+'.xlsx'
    reportpdf = str(select_doc)+'.pdf'
    reportpdfA = '/'+str(select_doc)+'.pdf'


# define the document to pull - bot sure what the back slash, but will find out 

    #print(report)
    #print(reportA)
    #print(reportpdf)
    #print(reportpdfA)

    ## Pull the document from dropbox in excel format, select # products sheet, convert to pdf, and put pack

    select_productno = str(NoProducts)

    with open(report, "wb") as f: # this is the desired name
        metadata, res = dbx.files_download(path=reportA) # this is the source file name
        f.write(res.content)
    
    import convertapi

    #convertapi.api_secret = 'fln0ZGtFJ9O7Iv9U'

    convertapi.api_secret = 'NF7ekuS1pbpX5cLH'

    result = convertapi.convert(
        'pdf',
        {
            'File': report,
            'WorksheetName': ''+(select_productno)+' Products (2)',
            'PdfResolution': '150',
        }
    )



    result.file.save(reportpdf)


    with open(reportpdf, "rb") as f: # this is the file just converted
        dbx.files_upload(f.read(), reportpdfA, mute = True, mode=dropbox.files.WriteMode.overwrite) # This is the desired name


    ## Rename the file, prep for e-mail circulation 

    import os
    os.rename(reportpdf,'Quotation2.pdf') # FIND MORE SECURE NAMING CONVENTION ID

#parameters "https://www.convertapi.com/xlsx-to-pdf#snippet=python"


def sent_updater_withOTP(sheet1,column):
    
    ################# Inumerate to confirm if we can send out e-mail #######################

    
    now = datetime.now()
    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")
    sheet = sheet1
    cell_list = Column_Length_OTP(column,sheet) # the columns of the "sent" reference
    sheet4 = cell_list[1]
    sent_no = len([i for i, s in enumerate(cell_list[0]) if 'sent' in s]) 
    tosend_no = len([i for i, s in enumerate(cell_list[0]) if 'sent' not in s])
    
    if tosend_no !=0:
        sent = [i for i, s in enumerate(cell_list[0]) if 'sent' in s]
        tosend = [i for i, s in enumerate(cell_list[0]) if 'sent' not in s]
        row_num = (tosend[0]+1 )#+ nottosend[0])
    
    
    #for val in tosend:
     #   if val == 0:
      #      break
        
        

    
    #print("Can proceed with generating output:")
    #Market_Report(row_num)
        #Dupont_Doc_Row(data_row) # - this is crux of ontodocs - the document creation function **CRUX-ONTODOCS**
     #   Yes_sendmail(Co_name,to_email) # still dont understand where it gets the e-mail from
      #  sheet4.update_cell(row_num, 16, 'sent A')
    
    #sheet4.update_cell(row_num, 4, 'sent Z')



    #print("Cell_List :",cell_list[0],len(cell_list[0]))
    #print("Sent No :",sent_no) # items where there is a "sent"
    #print("To Send No :",tosend_no) # items where there is a "sent"
    #print("To Send No :",tosend_no) # items where there is a "sent"
    #print("Next Row Number to update",row_num)
    data_row = sheet4.row_values(row_num)
    #Co_name = data_row[0] #- correct
    to_email = data_row[1]
    products_no = data_row[3] # cant be blank
    #print("The e-mail goes to:",to_email," for product(s) quantity:",products_no)
    
    
    
    ################# Function Pull the product for pdf, before sending e-mail #######################

    
    #NoProducts = 2
    Pull_LatestQuote_clean(products_no)

    
    ################# Function Pull the product for pdf, before sending e-mail #######################

    OTP = sheet4.cell(row_num, 10).value
    Yes_OTPsendmail(OTP,to_email)
    
    #Co_name = 'Bofin Consulting'
    
    #Yes_sendmail(Co_name,to_email)
    
    
    ################# Update the row that we have sent out the e-mail #######################

    sheet4.update_cell(row_num, (ord(column)-64), 'sent Quote') #16 is equivalent to column"P"
    #print("Succeessful: We have sent out an e-mail")

    
#column = "U"  ### Remember it must align to the lengh ie 21 
#sheet1 = "Quotations Manager"

#sent_updater_withOTP(sheet1,column)


    
secs = 15
def OTP_Runner_Silent():
    
    column = "U"
    sheet1 = "Quotations Manager"

    try:
        
        sent_updater_withOTP(sheet1,column)
        
        
    except:
        File_Name = 'Uptime Monitor Sheet'
        sheetA = "Uptime Monitor"
        Column_No = 4
        Column = "D"
        sheet_updater_silent(File_Name,sheetA,Column,Column_No)
        
schedule.every(secs).seconds.do(OTP_Runner_Silent)

while True:
    schedule.run_pending()
    time.sleep(1)



    
