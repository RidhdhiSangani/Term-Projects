import cv2
from cv2 import *
import numpy as np
import face_recognition
import os
from datetime import datetime
import pandas as pd
import smtplib
import ssl
from email.message import EmailMessage
import requests
from requests.exceptions import Timeout

path = '/Users/ridhdhi/Downloads/People'
images = []
classNames = []
myList = os.listdir(path)
print(myList)

for cl in myList:
    curImg = cv2.imread(f'{path}/{cl}')
    images.append(curImg)
    classNames.append(os.path.splitext(cl)[0])
print(classNames)

def findEncodings(images):
    encodeList = []
    for img in images:
        img = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
        encode = face_recognition.face_encodings(img)[0]
        encodeList.append(encode)
    return encodeList
def markAttendance(name):
    now = datetime.now()
    dtString = now.strftime('%d/%m/%Y')
    df=pd.read_excel('Attendance.xlsx',index_col=0)
    dt=df.columns
    # print(dt)
    names=df["Name"].values.tolist()
    # print(names)
    if(dt[len(dt)-1]!=dtString):
        df.insert(len(df.columns), dtString, 'A')
    df.at[names.index(name),dtString]='P'
    # print(df)
    df.to_excel('Attendance.xlsx')
def sendemail():
    # your_email = "attendancesystememail@gmail.com"
    # your_password = "ufwfedkwshfugdsw"
    your_email = "attendancesystememail@gmail.com"
    # your_password = "jaux zbnw ocnb uuzk"
    your_password = 'jrjhjnzaxvuiuaxo'
    # server = smtplib.SMTP_SSL('smtp.gmail.com', 587)
    # try:
    #     server = smtplib.SMTP_SSL('smtp.gmail.com',465)
    # except:
    #     print("Bad luck")

    # server = smtplib.SMTP('smtp.gmail.com', 587) 
    # server = smtplib.SMTP(localhost)
    # server = smtplib.SMTP('smtp.gmail.com:587')
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.set_debuglevel(1)
    server.ehlo()
    # server.strarttls()
    # server.ehlo()
    server.login(your_email, your_password)
    df=pd.read_excel('Attendance.xlsx',index_col=0)
    dt=df.columns
    print(dt)
    p=0
    a=0
    for i, r in df.iterrows():
        for j in range(2,len(dt)):
            if(r[dt[j]]=='P'):
                p+=1
            else :
                a+=1
        print("Your attendance is ", (str)(p/(p+a)*100))
        s="Your attendance is "+ (str)(p/(p+a)*100)+"percent"
        msg = EmailMessage()
        msg['Subject'] = "Your Attendance Report"
        msg['From'] = your_email
        msg['To'] = r["Email"]
        msg.set_content(s)
        server.send_message(msg)
    server.quit()
    print("Message sent")
def addstd():
    df=pd.read_excel('Attendance.xlsx',index_col=0)
    names=df["Name"].values.tolist()
    while True:
        print("Enter name")
        nm=input()
        nm=nm.upper()
        if nm in names:
            print("Name already present")
        else:
            break
    print("Enter email")
    em=input()
    l="People/"+nm+".jpg"
    cap = cv2.VideoCapture(0)
    while True:
        result, img = cap.read()
        cv2.imshow('Webcam',img)
        if (cv2.waitKey(1) & 0xFF == ord('c')): 
            if result:
                cv2.imwrite(l, img)
            cv2.destroyAllWindows()
            cv2.waitKey(1)
            break
    d={}
    dt=df.columns;
    d["Name"]=nm
    d["Email"]=em
    l=[]
    l.append(nm)
    l.append(em)
    for j in range(2,len(dt)):
        d[dt[j]]='A'
        l.append('A')
    print(d,l)
    df.loc[len(df)]=l
    print(df)
    df.to_excel('Attendance.xlsx')
    
encodeListKnown = findEncodings(images)
print('Encoding Complete')
while True:
    print("1.Take Attendance")
    print("2.Send mail")
    print("3.Add student")
    c=int(input())
    if(c==1):
        cap = cv2.VideoCapture(0)
        
        while True:
            success, img = cap.read()
            imgS = cv2.resize(img, (0,0), None, 0.25,0.25)
            imgS = cv2.cvtColor(imgS, cv2.COLOR_BGR2RGB)
        
            facesCurFrame = face_recognition.face_locations(imgS)
            encodesCurFrame = face_recognition.face_encodings(imgS, facesCurFrame)
        
            for encodeFace, faceLoc in zip(encodesCurFrame, facesCurFrame):
                matches = face_recognition.compare_faces(encodeListKnown, encodeFace)
                faceDis = face_recognition.face_distance(encodeListKnown, encodeFace)
                # print(faceDis)
                matchIndex = np.argmin(faceDis)
        
                if matches[matchIndex]:
                    name= classNames[matchIndex].upper()
                    # print(name)
                    y1,x2,y2,x1 = faceLoc
                    y1, x2, y2, x1 = y1*4,x2*4,y2*4,x1*4
                    cv2.rectangle(img, (x1,y1),(x2,y2),(0,255,0),2)
                    cv2.rectangle(img,(x1,y2-35),(x2,y2),(0,255,0),cv2.FILLED)
                    cv2.putText(img,name,(x1+6,y2-6),cv2.FONT_HERSHEY_COMPLEX,1,(255,255,255),2)
                    print(matches,faceDis)
                    markAttendance(name)
        
            cv2.imshow('Webcam',img)
            if (cv2.waitKey(1) & 0xFF == ord('q')):
                cap.release()

                cv2.destroyAllWindows()
                cv2.waitKey(1)
                break
    elif(c==2):
        sendemail()
    elif(c==3):
        addstd()
    else:
        break