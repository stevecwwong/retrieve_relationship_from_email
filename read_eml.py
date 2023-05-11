from pathlib import Path
import pandas as pd
import extract_msg
import email
import os
import xlsxwriter
import xml.etree.ElementTree as ET
import re
import text2emotion as te
import magic
from nrclex import NRCLex
from bs4 import BeautifulSoup

def df_get_name(email, df_user):
    try:
        name = df_user.loc[df_user['email'] == email, 'name'].item()
        return name
    except ValueError:
        return None  # Code not found in DataFrame

def df_retrieve_user(df_input):
    df_user = pd.DataFrame(columns = ['name', 'email'])
    i = 0
    
    for index, row in df_input.iterrows():
        cc_email = row['cc_email']
        if len(row['sender_email']) > 0 and len(row['sender_name']) > 0:
            #if row['sender_email'] not in df_user['email']:
            if not (df_user['email'] == row['sender_email']).any():
                sender = row['sender_name']
                sender0 = sender[0]
                if sender[0] == "" or sender[0] == " ":
                    sender0 = row['sender_email']
                    if "@" in sender0:
                        sender0 = sender0.split("@")[0]
                    else:
                        sender0 = sender0

                df_user.loc[len(df_user.index)] = [sender0.strip(), row['sender_email']]
        elif len(row['sender_email']) > 0 and len(row['sender_name']) == 0:
            if not (df_user['email'] == row['sender_email']).any():

                sender = row['sender_name']
                if len(sender) == 0 or sender == [""]:
                    
                    sender0 = row['sender_email']
                    if "@" in sender0:
                        sender0 = sender0.split("@")[0]
                    else:
                        sender0 = sender0
                df_user.loc[len(df_user.index)] = [sender0.strip(), row['sender_email']]

        if row['cc_email'] and not (df_user['email'] == cc_email[0]).any():
            sender0 = row['cc_name']
            if len(row['cc_name']) == 0:
                sender = row['cc_email']
                if "@" in sender[0]:
                    sender0 = sender[0].split("@")[0]
                else:
                    sender0 = sender[0]

            df_user.loc[len(df_user.index)] = [sender0, cc_email[0]]


    for index, row in df_input.iterrows():
        i = 0
        to_email = row['to_email']
        to_name = row['to_name']
        for i in range(len(to_email)):
            
            # print("i:",index, i, to_email,  " index:", index, to_name, len(to_name))
            # if len(to_name) == 0:
            #     to_name_i = ""
            # elif len(to_name) < i+1:
            #     to_name_i = ""
            # elif len(to_name) >= i+1:
            #     to_name_i = to_name[i]
            # else:
            #     to_name_i = ""
            # print("ss:", to_name_i)

            # print("1len:",index, len(sender),i, to_name, row['to_name'],to_email[i])

            if row['to_email'] and not (df_user['email'] == to_email[i]).any() :
                if "@" in to_email[i]:
                    sender0 = to_email[i].split("@")[0]
                else:
                    sender0 = sender[0]

                if len(sender) >= i:
                    df_user.loc[len(df_user.index)] = [sender0, to_email[i]]
                else:
                    df_user.loc[len(df_user.index)] = [sender0, to_email[i]]

    df_user.loc[len(df_user.index)] = ["Mr Evil", "whoknowsme@sbcglobal.net"]
    return df_user

def df_retrieve_relation(df_input, df_user):

    
    to_e = ""
    to_n = ""
    df_relation = pd.DataFrame(columns = ['image', 'subject', 'source', 'source_name', 'target', 'target_name', 'dates', 'body', 'emotion', 'nrclex', 'happy', 'angry', 'surprise', 'sad', 'fear', 'nfear', 'nanger', 'nanticipation', 'ntrust', 'nsurprise', 'npositive', 'nnegative', 'nsadness', 'ndisgust', 'njoy'])
#    print("rowcount:",df_input.shape[0])
    i = 0
    for index, row in df_input.iterrows():
        if row['image'] == "4dell":
            default_to_e = "whoknowsme@sbcglobal.net"
            default_to_n = "Mr Evil"
        elif row['image'] == "mantooth32":
            default_to_e = "dollarhyde86@comcast.net"
            default_to_n = "Wes Mantooth"
        elif row['image'] == "washer":
            default_to_e = "chkwasher@comcast.net"
            default_to_n = "John Washer"
        else:
            default_to_e = ""
            default_to_n = ""
#        print("index:",index)
#        print("from:",row['sender_email'])
        to_e = row['to_email'] + row['cc_email']
        to_n = row['to_name']
        save_e = ""
        save_n = ""
        save_image = ""
        # print("to:",row['to_email'], len(row['to_email']))
        # print("cc:",row['cc_email'], len(row['cc_email']))
        i = 0
        if row['sender_email'] == [] and len(to_e) == 0:
            continue
        elif len(to_e) == 0:
            lookup_sender_name = df_get_name(row['sender_email'], df_user)
            print("sender:", lookup_sender_name, row['sender_email'])
            lookup_target_name = df_get_name(default_to_e, df_user)
            print("receiver:", lookup_target_name, default_to_e)
            

            #df_get_name(row['default_to_e'], df_user)
            df_relation.loc[len(df_relation.index)] = [row['image'], row['subject'], row['sender_email'], lookup_sender_name, default_to_e, lookup_target_name, row['dates'], row['body'], row['emotion'], row['nrclex'], row['happy'], row['angry'], row['surprise'], row['sad'], row['fear'], row['nfear'], row['nanger'], row['nanticipation'], row['ntrust'], row['nsurprise'], row['npositive'], row['nnegative'], row['nsadness'], row['ndisgust'], row['njoy']]
        else:
            for i in range(len(to_e)):
                send_email = row['sender_email']

                if row['sender_email'] == []:
                    send_email = default_to_e

                if len(to_e[i]) == 0 and row['image'] == '4dell':
                    save_e = default_to_e
                else:
                    save_e = to_e[i]
            
                lookup_sender_name = df_get_name(send_email, df_user)
                print("b sender:", lookup_sender_name, send_email)
                lookup_target_name = df_get_name(save_e, df_user)
                print("b receiver:", lookup_target_name, save_e)
            
#                print("row:",row['image'], " sender:",row['sender_email'])
                df_relation.loc[len(df_relation.index)] = [row['image'], row['subject'], send_email,  lookup_sender_name, save_e, lookup_target_name, row['dates'], row['body'], row['emotion'], row['nrclex'], row['happy'], row['angry'], row['surprise'], row['sad'], row['fear'], row['nfear'], row['nanger'], row['nanticipation'], row['ntrust'], row['nsurprise'], row['npositive'], row['nnegative'], row['nsadness'], row['ndisgust'], row['njoy']]

#    print("df_Relation length:", len(df_relation))

    return df_relation


# purpose - to resolve email string for name and email
def get_email_info(text):
    text = text.replace('"','').replace("'","")
    
    emails = re.findall(r'[\w.+-]+@[\w-]+\.[\w.-]+', text)
    pattern = r'([\w\-\ \.]+)\s*<([\w\.-]+@[\w\.-]+)>'
    match = re.search(pattern, text)
    matchs = re.findall(pattern, text)
    if len(matchs) > 0:
        se_name = [i[0] for i in matchs]
    else:
        if match:
            se_name[0] = match.group(1)
            email = match.group(2)
        else:
            se_name = []

    return se_name, emails

# purpose - open html email file
def get_html_info(file_path,image):
    try: 

        filetype = magic.from_file(file_path)
        with open(file_path, 'r') as f:
            content = f.readlines()
            for line in content:
                if "From:" in line:
                    sender = line.split("<TD>")[1].split("</TD>")[0]
                    break
            for line in content:
                if "To:" in line:
                    to = line.split("<TD>")[1].split("</TD>")[0]
                    break
            
            for line in content:
                if "Sent:" in line:
                    sent_address = line.split("<TD>")[1].split("</TD>")[0]
                    break
            
            for line in content:
                if "Subject:" in line:
                    subject = line.split("<TD>")[1].split("</TD>")[0]
                    break
            for line in content:
                if "Sent:" in line:
                    dates = line.split("<TD>")[1].split("</TD>")[0]
                    break

        if sender is not None:
            sender_info = get_email_info(str(sender))
#            print("sender_info:",sender_info)
            if len(sender_info) >= 2 and len(sender_info[1]) >= 1:# and not sender_info:
                s_email = sender_info[1][0]
                s_name = sender_info[0]
            else:
                s_email = []
                s_name = []
        else:
            s_email = []
            s_name = []
        
        if to is not None:
            to_info = get_email_info(str(to))
#            print("to_info:",to_info)
            if len(to_info) >= 2 and len(to_info[1]) >= 1:# and not sender_info:
                t_email = to_info[1]
                t_name = to_info[0]
            else:
                t_email = []
                t_name = []
        else:
            t_email = []
            t_name = []

        c_name = []
        c_email = []

        with open(file_path, 'r') as f:
            email_string = f.read()
            msg = email.message_from_string(email_string)
            start = str(msg).find('<TABLE cellspacing=0 class="emlbdy">')
            end = str(msg).find('</TBODY></TABLE>',start)
            body_str = str(msg)
            body = body_str[start:end]
            body_loc = body.find("<".encode('unicode_escape').decode())
            #print("body_loc:",body_loc)
            while body_loc != -1:
                start1 = body.find("<".encode('unicode_escape').decode())
                end1 = body.find(">".encode('unicode_escape').decode())
                result = body[start1:end1+1]
                #print("result:",result)
                body = body.replace(result,"")
                body_loc = body.find("<".encode('unicode_escape').decode())
                #print("body replace:",body)

            t_happy = 0
            t_angry = 0
            t_surprise = 0
            t_sad = 0
            t_fear = 0
            n_fear = 0
            n_anger = 0
            n_anticipation = 0
            n_trust = 0
            n_surprise = 0
            n_positive = 0
            n_negative = 0
            n_sadness = 0
            n_disgust = 0
            n_joy = 0

            emotion = {}
            nrc_emotion = {}
            nrc_return = {}
            if len(body) >= 1:
                emotion = te.get_emotion(body)
                t_happy = emotion.get('Happy')
                t_angry = emotion.get('Angry')
                t_surprise = emotion.get('Surprise')
                t_sad = emotion.get('Sad')
                t_fear = emotion.get('Fear')
                nrc_emotion = NRCLex(body)
                nrc_return= nrc_emotion.affect_frequencies
                n_fear = nrc_return.get('fear')
                n_anger = nrc_return.get('anger')
                n_anticipation = nrc_return.get('anticipation')
                n_trust = nrc_return.get('trust')
                n_surprise = nrc_return.get('surprise')
                n_positive = nrc_return.get('positive')
                n_negative = nrc_return.get('negative')
                n_sadness = nrc_return.get('sadness')
                n_disgust = nrc_return.get('disgust')
                n_joy = nrc_return.get('joy')

    except Exception as e: 
        print("get_xml_info",file_path, e)

    return {'image': image, 'type': filetype, 'filename': file_path, 'subject': subject, 'sender': sender, 'sender_name': s_name, 'sender_email': s_email, 'to': to, 'to_name': t_name, 'to_email': t_email, 'cc': cc, 'cc_name': c_name, 'cc_email': c_email, 'dates': dates, 'body': body, 'emotion': emotion, 'nrclex': nrc_return, 'happy': t_happy, 'angry': t_angry, 'surprise': t_surprise, 'sad': t_sad, 'fear': t_fear, 'nfear': n_fear, 'nanger': n_anger, 'nanticipation' : n_anticipation, 'ntrust': n_trust, 'nsurprise' : n_surprise, 'npositive' : n_positive, 'nnegative' : n_negative, 'nsadness': n_sadness, 'ndisgust' : n_disgust, 'njoy': n_joy}
    #return {'image': image, 'type': filetype, 'filename': file_path}

# purpose - check if file is xml or html, or not
def parse_report_file(report_input_file):
    with open(report_input_file) as unknown_file:
        c = unknown_file.read(1)
        if c != '<':
            return 'Is JSON'
        return 'Is XML'


def get_file_list(directory):
    os.chdir(directory)
    return [f for f in os.listdir(directory) if os.path.isfile(os.path.join("", f))]

def get_eml_info(file_path,image):
    try: 
        filetype = magic.from_file(file_path)
        with open(file_path, 'rb') as f:
            msg = email.message_from_binary_file(f)
        
        subject = msg['subject']
        sender = msg['from']
        to = msg['to']
        dates = msg['date']
        cc = msg['cc']
        body = None
        if sender is not None:
            sender_info = get_email_info(str(sender))
#            print("sender_info:",sender_info)
            if len(sender_info) >= 2 and len(sender_info[1]) >= 1:# and not sender_info:
                s_email = sender_info[1][0]
                s_name = sender_info[0]
            else:
                s_email = []
                s_name = []
        else:
            s_email = []
            s_name = []
        
        if to is not None:
            to_info = get_email_info(str(to))
#            print("to_info:",to_info)
            if len(to_info) >= 2 and len(to_info[1]) >= 1:# and not sender_info:
                t_email = to_info[1]
                t_name = to_info[0]
            else:
                t_email = []
                t_name = []
        else:
            t_email = []
            t_name = []

        if cc is not None:
            cc_info = get_email_info(str(cc))
#            print("to_info:",cc_info)
            if len(cc_info) >= 2 and len(cc_info[1]) >= 1:# and not sender_info:
                c_email = cc_info[1]
                c_name = cc_info[0]
            else:
                c_email = []
                c_name = []
        else:
            c_email = []
            c_name = []
        
        #to_email = get_email_info(to)
#        print("sender:",sender_name, " email:",s_email)
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                content_charset = part.get_content_charset()
#                print("if:",file_path,content_type,content_charset)
                if content_type == 'text/plain':
                    if content_charset != None:
                        body = part.get_payload(decode=True).decode(content_charset)
                        break
                    else:                        
                        body = part.get_payload(decode=True).decode()
                        break
        else:
            content_charset = msg.get_content_charset()
            if content_charset == 'iso-1252':
#                print("else:", file_path, content_charset)
                body = msg.get_payload(decode=True).decode()
            elif content_charset != None:
#                print("else2:", file_path, content_charset)
                body = msg.get_payload(decode=True).decode(content_charset)
#                print("body:", body)
            else:
#                print("else2:", file_path, content_charset)
                body = msg.get_payload(decode=True).decode()
        
        # if isinstance(body,list):
        #     print("body is list")
        # else:
        #     print("bosy is char")
        t_happy = 0
        t_angry = 0
        t_surprise = 0
        t_sad = 0
        t_fear = 0
        n_fear = 0
        n_anger = 0
        n_anticipation = 0
        n_trust = 0
        n_surprise = 0
        n_positive = 0
        n_negative = 0
        n_sadness = 0
        n_disgust = 0
        n_joy = 0
        emotion = {}
        nrc_emotion = {}
        nrc_return = {}
        if len(body) >= 1:
            emotion = te.get_emotion(body)
            print("text2emotion:",emotion) #happy, angry,surprise,sad,fear)
            t_happy = emotion.get('Happy')
            t_angry = emotion.get('Angry')
            t_surprise = emotion.get('Surprise')
            t_sad = emotion.get('Sad')
            t_fear = emotion.get('Fear')
            nrc_emotion = NRCLex(body)
            nrc_return = nrc_emotion.affect_frequencies
            print("nrclex:",nrc_return) 
            n_fear = nrc_return.get('fear')
            n_anger = nrc_return.get('anger')
            n_anticipation = nrc_return.get('anticipation')
            n_trust = nrc_return.get('trust')
            n_surprise = nrc_return.get('surprise')
            n_positive = nrc_return.get('positive')
            n_negative = nrc_return.get('negative')
            n_sadness = nrc_return.get('sadness')
            n_disgust = nrc_return.get('disgust')
            n_joy = nrc_return.get('joy')

    except Exception as e: 
        print("get_eml_info",file_path, e)

#    return {'image': image, 'filename': file_path, 'subject': subject, 'sender': sender, 'sender_name': sender_name, 'sender_email': sender_email, 'to': to, 'to_name': to, 'to_email': to_email, 'dates': dates, 'body': body}
    return {'image': image, 'type': filetype, 'filename': file_path, 'subject': subject, 'sender': sender, 'sender_name': s_name, 'sender_email': s_email, 'to': to, 'to_name': t_name, 'to_email': t_email, 'cc': cc, 'cc_name': c_name, 'cc_email': c_email, 'dates': dates, 'body': body, 'emotion': emotion, 'nrclex': nrc_return, 'happy': t_happy, 'angry': t_angry, 'surprise': t_surprise, 'sad': t_sad, 'fear': t_fear, 'nfear': n_fear, 'nanger': n_anger, 'nanticipation' : n_anticipation, 'ntrust': n_trust, 'nsurprise' : n_surprise, 'npositive' : n_positive, 'nnegative' : n_negative, 'nsadness': n_sadness, 'ndisgust' : n_disgust, 'njoy': n_joy}

def save_email_to_df(df,name,path):
    filepath = get_file_list(path)
    for file in filepath:
        if parse_report_file(file) == 'Is XML':
#            print("xml:",file)
            new_info = get_html_info(file,name)
            df.loc[len(df)] = new_info
        else:
            new_info = get_eml_info(file,name)
            df.loc[len(df)] = new_info
    return df

try:
    image_files = {"name":["4dell","mantooth32","washer"], "path":["C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\email\\4dell","C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\email\\mantooth32","C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\email\\washer"]}

    senders = []
    dates = []
    subjects = []
    bodies = []
    to = []
    to_email = []
    to_name = []
    sender_name = []
    sender_email = []
    cc = []
    df_eml = pd.DataFrame(columns = ['image', 'type', 'filename', 'subject', 'sender', 'sender_name','sender_email', 'to', 'to_name', 'to_email', 'cc', 'cc_name', 'cc_email', 'dates', 'body', 'emotion', 'nrclex', 'happy', 'angry', 'surprise', 'sad', 'fear', 'nfear', 'nanger', 'nanticipation', 'ntrust', 'nsurprise', 'npositive', 'nnegative', 'nsadness', 'ndisgust', 'njoy'])

    for i in range(len(image_files)+1):
        df_eml = save_email_to_df(df_eml,image_files['name'][i], image_files['path'][i])

    df_user = df_retrieve_user(df_eml)
    df_relation = df_retrieve_relation(df_eml, df_user)

#print(df_eml)
    writer = pd.ExcelWriter(r'C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\document\\emaillist.xlsx',options={'strings_to_urls': False})
    df_eml.to_excel(writer)
    writer.close()

    writer_relation = pd.ExcelWriter(r'C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\document\\relation.xlsx',options={'strings_to_urls': False})
    df_relation.to_excel(writer_relation)
    writer_relation.close()

    writer_user = pd.ExcelWriter(r'C:\\Users\\Sam Cheng\\Desktop\\Steve\\bcit\\image\\case\\document\\user.xlsx',options={'strings_to_urls': False})
    df_user.to_excel(writer_user)
    writer_user.close()

except Exception as e: 
    print(e)