from ast import Break
import tkinter as tk
from tkinter import Message ,Text
import cv2,os
import shutil
import csv
import numpy as np
from PIL import Image, ImageTk
import pandas as pd
import datetime
import time
import tkinter.ttk as ttk
import tkinter.font as font
from csv import writer,DictWriter
windo = tk.Tk()
windo.attributes('-fullscreen',True)
windo.title("Attendance")


windo.configure(background='black')



windo.grid_rowconfigure(0, weight=1)
windo.grid_columnconfigure(0, weight=1) 
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
 
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
 
    return False 
def start():
    recognizer = cv2.face.LBPHFaceRecognizer_create()
    recognizer.read("Trainer.yml")
    harcascadePath = "haarcascade_frontalface_default.xml"
    faceCascade = cv2.CascadeClassifier(harcascadePath);    
    df=pd.read_csv("EmployeeDetails\EmployeeDetails.csv")
    cam = cv2.VideoCapture(0)
    font = cv2.FONT_HERSHEY_SIMPLEX        
    col_names =  ['Id','Name','Date','Time']
    attendance = pd.DataFrame(columns = col_names)    
    while True:
        ret, ic =cam.read()
        gray=cv2.cvtColor(ic,cv2.COLOR_BGR2GRAY)
        faces=faceCascade.detectMultiScale(gray, 1.2,5)    
        for(x,y,w,h) in faces:
            cv2.rectangle(ic,(x,y),(x+w,y+h),(225,0,0),2)
            Id, conf = recognizer.predict(gray[y:y+h,x:x+w])                                   
            if(conf < 50):
                ts = time.time()      
                date = datetime.datetime.fromtimestamp(ts).strftime("%d/%m/%Y")
                timeStamp = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
                aa=df.loc[df['Id'] == Id]['Name'].values
                tt="Hi "+str(Id)+" "+aa
                attendance.loc[len(attendance)] = [Id,aa,date,timeStamp]
                
            else:
                cam.release()
                cv2.destroyAllWindows()
                """root = tk.Tk()
                root.title('Yes/No')
                root.geometry('300x150')
                def yes():
                    newmember()
                def No():
                    root.destroy
                    start()
                yesbutton= tk.Button(root, text="Yes", command=yes)
                yesbutton.place(x=100,y=300)
                nobutton= tk.Button(root, text="No", command=No)
                nobutton.place(x=150,y=300)
                root.mainloop()""" 
                newmember()
                start
            cv2.putText(ic,str(tt),(x,y+h), font, 1,(255,255,255),2)        
        attendance=attendance.drop_duplicates(subset=['Id'],keep='first')
            
        cv2.imshow('Frame',ic) 
        if (cv2.waitKey(1)==ord('q')):
            cam.release()
            windo.destroy()
            break
    ts = time.time()      
    date = datetime.datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
    timeStamp = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
    Hour,Minute,Second=timeStamp.split(":")
    print(attendance)
    #print(date,timeStamp)
    path="Attendance\Attendance"+"_"+date+".xlsx"
    print(os.path.isfile(path))
    if(os.path.isfile(path)):
        filenam=os.path.basename(path)
        print(os.path.basename(path))
        df=pd.read_excel(path)
        print(df)
        df2=pd.concat([df,attendance],ignore_index=True)
        print(df2)
        #with open(filename,'a',encoding='UTF8') as f:
        fileName="Attendance\Attendance"+"_"+date+".xlsx"
        df2.to_excel(fileName,index=False)
            
    else:
        fileName="Attendance\Attendance"+"_"+date+".xlsx"
        attendance.to_excel(fileName,index=False) 
def newmember():
    window = tk.Tk()
    window.title("Attendance")

    window.attributes('-fullscreen',True)
    window.configure(background='black')



    window.grid_rowconfigure(0, weight=1)
    window.grid_columnconfigure(0, weight=1)


    message = tk.Label(window, text="Attendance Marking" ,bg="black"  ,fg="white"  ,width=50  ,height=3,font=('times', 30, 'italic bold underline')) 

    message.place(x=200, y=20)

    lbl = tk.Label(window, text="Enter ID",width=20  ,fg="red"  ,bg="white" ,font=('times', 15, ' bold ') ) 
    lbl.place(x=400, y=215)

    txt = tk.Entry(window,width=20,bg="white" ,fg="red",font=('times', 15, ' bold '))
    txt.place(x=700, y=215)

    lbl2 = tk.Label(window, text="Enter Name",width=20  ,fg="red"  ,bg="white",height=1 ,font=('times', 15, ' bold ')) 
    lbl2.place(x=400, y=315)

    txt2 = tk.Entry(window,width=20,bg="white"  ,fg="red",font=('times', 15, ' bold ')  )
    txt2.place(x=700, y=315)    
    f=0;
    def clear():
        txt.delete(0, 'end')
        txt2.delete(0,'end')
    def TakeImages():  
        Id=(txt.get())
        name=(txt2.get())
        if(is_number(Id) and name.isalpha()):
            cam = cv2.VideoCapture(0)
            harcascadePath = "haarcascade_frontalface_default.xml"
            detector=cv2.CascadeClassifier(harcascadePath)
            sampleNum=0
            while(True):
                ret, img = cam.read()
                gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
                faces = detector.detectMultiScale(gray, 1.3, 5)
                for (x,y,w,h) in faces:
                    cv2.rectangle(img,(x,y),(x+w,y+h),(255,0,0),2)        
                    sampleNum=sampleNum+1
                    cv2.imwrite("TrainingImage\ "+name +"."+Id +'.'+ str(sampleNum) + ".jpg", gray[y:y+h,x:x+w])
                    cv2.imshow('frame',img) 
                if cv2.waitKey(100) & 0xFF == ord('q'):
                    break
                elif sampleNum>60:
                    cv2.destroyAllWindows()
                    break
            cam.release()
            res = "Images Saved for ID : " + Id +" Name : "+ name
            row = [Id , name]
            with open('EmployeeDetails\EmployeeDetails.csv','a+') as csvFile:
                writer = csv.writer(csvFile)
                writer.writerow(row)
            csvFile.close()
        def getImagesAndLabels(path):
            imagePaths=[os.path.join(path,f) for f in os.listdir(path)] 
            #print(imagePaths)
            faces=[]
            Ids=[]
            for imagePath in imagePaths:
                pilImage=Image.open(imagePath).convert('L')
                imageNp=np.array(pilImage,'uint8')
                Id=int(os.path.split(imagePath)[-1].split(".")[1])
                faces.append(imageNp)
                Ids.append(Id)        
            return faces,Ids
        recognizer = cv2.face.LBPHFaceRecognizer.create()
        harcascadePath = "haarcascade_frontalface_default.xml"
        detector =cv2.CascadeClassifier(harcascadePath)
        faces,Id = getImagesAndLabels("TrainingImage")
        recognizer.train(faces, np.array(Id))
        recognizer.save("Trainer.yml")
        cam.release()
        window.destroy()
        start()
        recognizer = cv2.face.LBPHFaceRecognizer_create()
        recognizer.read("Trainer.yml")
        harcascadePath = "haarcascade_frontalface_default.xml"
        faceCascade = cv2.CascadeClassifier(harcascadePath);    
        df=pd.read_csv("EmployeeDetails\EmployeeDetails.csv")
        cam = cv2.VideoCapture(0)
        font = cv2.FONT_HERSHEY_SIMPLEX        
        col_names =  ['Id','Name','Date','Time']
        attendance = pd.DataFrame(columns = col_names)    
        while True:
            ret, im =cam.read()
            gray=cv2.cvtColor(im,cv2.COLOR_BGR2GRAY)
            faces=faceCascade.detectMultiScale(gray, 1.2,5)    
            for(x,y,w,h) in faces:
                cv2.rectangle(im,(x,y),(x+w,y+h),(225,0,0),2)
                Id, conf = recognizer.predict(gray[y:y+h,x:x+w])                                   
                if(conf < 50):
                    ts = time.time()      
                    date = datetime.datetime.fromtimestamp(ts).strftime("%d/%m/%Y")
                    timeStamp = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
                    aa=df.loc[df['Id'] == Id]['Name'].values
                    tt="Hi "+str(Id)+" "+aa
                    attendance.loc[len(attendance)] = [Id,aa,date,timeStamp]     
                cv2.putText(im,str(tt),(x,y+h), font, 1,(255,255,255),2)        
            attendance=attendance.drop_duplicates(subset=['Id'],keep='first')
                
            cv2.imshow('im',im) 
            if (cv2.waitKey(1)==ord('q')):
                cam.release()
                cv2.destroyAllWindows()
                break
        ts = time.time()      
        date = datetime.datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timeStamp = datetime.datetime.fromtimestamp(ts).strftime('%H:%M:%S')
        Hour,Minute,Second=timeStamp.split(":")
        print(attendance)
        #print(date,timeStamp)
        path="Attendance\Attendance"+"_"+date+".xlsx"
        print(os.path.isfile(path))
        if(os.path.isfile(path)):
            filenam=os.path.basename(path)
            #print(os.path.basename(path))
            df=pd.read_excel(path)
            #print(df)
            df2=pd.concat([df,attendance],ignore_index=True)
            print(df2)
            #with open(filename,'a',encoding='UTF8') as f:
            fileName="Attendance\Attendance"+"_"+date+".xlsx"
            df2.to_excel(fileName,index=False)
            print("Done")
            window.destroy()
        else:
            fileName="Attendance\Attendance"+"_"+date+".xlsx"
            attendance.to_excel(fileName,index=False)
            print("Done")
            window.destroy()
    takeImg = tk.Button(window, text="Add New", command=TakeImages  ,fg="red"  ,bg="white"  ,width=20  , activebackground = "Red" ,font=('times', 15, ' bold '))
    takeImg.place(x=400, y=500)
    quitWindow = tk.Button(window, text="Quit", command=window.destroy  ,fg="red"  ,bg="white"  ,width=20 , activebackground = "Red" ,font=('times', 15, ' bold '))
    quitWindow.place(x=1000, y=500)
    copyWrite = tk.Text(window, background=window.cget("background"), borderwidth=0,font=('times', 30, 'italic bold underline'))
    copyWrite.tag_configure("superscript", offset=10)
    copyWrite.configure(state="disabled",fg="red"  )
    copyWrite.pack(side="left")
    copyWrite.place(x=800, y=750)
    window.mainloop()
  
StartButton = tk.Button(windo, text="Start", command=start  ,fg="red"  ,bg="white"  ,width=20  ,height=2 ,activebackground = "Red" ,font=('times', 15, ' bold '))
StartButton.place(x=950, y=200)
quitWindow = tk.Button(windo, text="Quit", command=windo.destroy  ,fg="red"  ,bg="white"  ,width=20  ,height=2, activebackground = "Red" ,font=('times', 15, ' bold '))
quitWindow.place(x=950, y=500)
windo.mainloop()