import getpass
import pandas as pd
import datetime as dt
from tkinter import *
import tkinter as tk
from tkinter import ttk
from PIL import Image,ImageTk
import smtplib

EMAIL_ID= 'ArhamOg@outlook.com' # my gmail 

# password from page
def get_password():
    password_window = Toplevel(root)
    password_window.title("Enter Password")
    password_window.geometry("300x100")

    password_label = Label(password_window, text="Enter your email password:")
    password_label.pack(pady=5)
    password_entry = Entry(password_window, show="*")
    password_entry.pack(pady=5)

    def submit_password():
        global EMAIL_PSWD
        EMAIL_PSWD = password_entry.get() 
        password_window.destroy()

    submit_button = Button(password_window, text="Submit", command=submit_password)
    submit_button.pack(pady=5)

root = Tk()
root.title("Email Login")
root.geometry("200x200")

password_button = Button(root, text="Enter Password", command=get_password)
password_button.pack(pady=20)

root.mainloop()

# student sheet
def createExcel(insert=None):
    StudentData = {"name": ['Arham','Swamil','Kartik','Kushu','Seema','Sania','Jenny','Lucifer','Maham'],
                   "Prn": [1032220940,1032220941,1032220945,1032220948,1032220950,1032220951,1032220980,1032220856,1032220450],
                   "Email": ['sanjusayeed81@outlook.com','arhamtemp@outlook.com','arhamt3@outlook.com',
                             'arhamt4@outlook.com','sanjusayeed81@outlook.com', 'arhamt2@outlook.com',
                             'arhamt3@outlook.com','arhamt1@myyahoo.com','orwner@myyahoo.com'],
                    "Attendence": [80,90,65,79,35,71,84,95,55],
                    "CGPA" : [8.1,6.9,7.3,6.9,9.6,9.1,8.3,7.1,6.0]}
                             
    df= pd.DataFrame(StudentData)
    df = df.rename_axis("Serial No").reset_index()
    df.to_csv('StudList.csv',index=False)
    
    return df

# holiday list
def holidayList():
    HList= {"Holiday": ['Republic Day','ShivJayanti','Holi','Eid-ul-Fitr',
                        'Maharashtra-Day','Independence-Day','Rakshabandhan','Krishna-Janmashtami',
                        'Ganesh-Chaturthi','Navratri','Diwali-start,','Diwali-end','Christmas'],
            "date": ['26-01-24','19-02,24','25-03-24','11-04-24','01-05-24','15-08-24','19-08-24',
                     '26-08-24','07-09-24','12-10-24','28-10-24','04-11-24','25-12-24'],
            "Day": ['Friday','Monday','Monday','Thursday','Wednesday','Thursday','Monday',
                    'Monday','saturday','saturday','Monday','Monday','Wednesday']}
    
    hl= pd.DataFrame(HList)
    hl= hl.rename_axis("Serial No").reset_index()
    hl.to_csv('HolidayList.csv',index=False)
  
    return hl

# event list
def EventList():
    Elist= {"Event": ['Vasundhara','Hack-Mitwpu','Aarohan','GDSC-WOW','SIH-Pune','RIDE-MITWPU','Technovision'],
            "Start-date":['22-01-2024','19-03-2024','26-03-2024','13-04-2024','15-09-2024','01-10-2024','29-10-2024'],
            "End-Date":['25-01-2024','21-03-2024','28-03-2024','15-04-2024','18-09-2024','05-10-2024','31-10-2024']}
    
   
    EL= pd.DataFrame(Elist)
    EL=EL.rename_axis("Serial No").reset_index()
    EL.to_csv('EventList.csv',index=False)
   
    return EL

# exam schedule
def ExamSchedule():
    Exam={"Examination":['Mid-term sem(2,4,6,8)','Practical-Exam sem(2,4,6,8)','End-Term sem(2,4,6,8)','Mid-term sem(1,3,5,7)','Practical-Exam sem(1,3,5,7)','End-Term sem(1,3,5,7)'],
          "Start-date":['01-05-2024','26-04-2024','16-05-2024','15-10-2024','26-11-2024','16-12-2024'],
           "End-Date":['05-05-2024','03-05-2024','28-05-2024','20-10-2024','03-12-2024','28-12-2024']}
    
    ES= pd.DataFrame(Exam)
    ES=ES.rename_axis("Serial No").reset_index()
    ES.to_csv('ExamSchedule.csv',index=False)

    return ES

# all data in gui
def show_student_data():
    student_window = Toplevel(root)
    student_window.title("Student Data")
    student_window.geometry("1000x400")
    student_window.configure(bg="green")
    
    title_label = Label(student_window, text="Student Data", font=("Roboto", 32), bg="green", fg="white")
    title_label.pack(pady=(5, 0))
    tree_frame = Frame(student_window, bg="green")
    tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    style = ttk.Style(student_window)
    style.configure("Treeview", background="deepskyblue", foreground="black", fieldbackground="deepskyblue")
    tree = ttk.Treeview(student_window, columns=list(student_df.columns), show="headings", style="Treeview")
   
    for column in student_df.columns:
        tree.heading(column, text=column)
    for index, row in student_df.iterrows():
        tree.insert("", "end", values=list(row))
    
    scrollbar_y = ttk.Scrollbar(student_window, orient="vertical", command=tree.yview)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x = ttk.Scrollbar(student_window, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")

    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    tree.pack(expand=True, fill="both")
 
def show_upcoming_holidays():
    holiday_window = Toplevel(root)
    holiday_window.title("Upcoming Holidays")
    holiday_window.geometry("1000x400")
    holiday_window.configure(bg="green")
 
    title_label = Label(holiday_window, text="Holiday List", font=("Roboto", 32), bg="green", fg="white")
    title_label.pack(pady=(5, 0))
    tree_frame = Frame(holiday_window, bg="green")
    tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    style = ttk.Style(holiday_window)
    style.configure("Treeview", background="lightcoral", foreground="black", fieldbackground="lightcoral")
    tree = ttk.Treeview(holiday_window, columns=list(holiday_df.columns), show="headings", style="Treeview")
   
    for column in holiday_df.columns:
        tree.heading(column, text=column)
    for index, row in holiday_df.iterrows():
        tree.insert("", "end", values=list(row))
    
    scrollbar_y = ttk.Scrollbar(holiday_window, orient="vertical", command=tree.yview)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x = ttk.Scrollbar(holiday_window, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")

    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    tree.pack(expand=True, fill="both")


def show_upcoming_events():
    event_window = Toplevel(root)
    event_window.title("Upcoming Events")
    event_window.geometry("1000x400")
    event_window.configure(bg="green")

    title_label = Label(event_window, text="Upcoming Events", font=("Roboto", 32), bg="green", fg="white")
    title_label.pack(pady=(5, 0))
    tree_frame = Frame(event_window, bg="green")
    tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    style = ttk.Style(event_window)
    style.configure("Treeview", background="yellow", foreground="black", fieldbackground="yellow")
    tree = ttk.Treeview(event_window, columns=list(event_df.columns), show="headings", style="Treeview")

    for column in event_df.columns:
        tree.heading(column, text=column)
    for index, row in event_df.iterrows():
        tree.insert("", "end", values=list(row))
    
    scrollbar_y = ttk.Scrollbar(event_window, orient="vertical", command=tree.yview)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x = ttk.Scrollbar(event_window, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")

    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    tree.pack(expand=True, fill="both")

def show_exam_schedule():
    exam_window = Toplevel(root)
    exam_window.title("Exam Schedule")
    exam_window.geometry("1000x400")
    exam_window.configure(bg="green") 

    title_label = Label(exam_window, text="Exam Schedule", font=("Roboto", 32), bg="green", fg="white")
    title_label.pack(pady=(5, 0))
    tree_frame = Frame(exam_window, bg="green")
    tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    style = ttk.Style(exam_window)
    style.configure("Treeview", background="lightblue", foreground="black", fieldbackground="lightblue")
    tree = ttk.Treeview(exam_window, columns=list(exam_df.columns), show="headings", style="Treeview")

    for column in exam_df.columns:
        tree.heading(column, text=column)
    for index, row in exam_df.iterrows():
        tree.insert("", "end", values=list(row))
    
    scrollbar_y = ttk.Scrollbar(exam_window, orient="vertical", command=tree.yview)
    scrollbar_y.pack(side="right", fill="y")
    scrollbar_x = ttk.Scrollbar(exam_window, orient="horizontal", command=tree.xview)
    scrollbar_x.pack(side="bottom", fill="x")
    
    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    tree.pack(expand=True, fill="both")



def display_data():    
    global student_df,exam_df,event_df,holiday_df,root
    student_df = pd.read_csv('StudList.csv')
    holiday_df = pd.read_csv('HolidayList.csv')
    event_df = pd.read_csv('EventList.csv')
    exam_df = pd.read_csv('ExamSchedule.csv')

    root=Tk()
    root.title("University Update Info")
    root.geometry("1000x1000")
    root.configure(bg="turquoise")
    
    logo_image= Image.open("images.jpeg")
    logo_photo= ImageTk.PhotoImage(logo_image)
    root.grid_rowconfigure(0, pad=10)
    root.grid_columnconfigure(0, pad=10)
    
    side_frame = Frame(root, width=50, bg="navy")
    side_frame.pack(side="left", fill="y")
    side_frame = Frame(root, width=50, bg="navy")
    side_frame.pack(side="right", fill="y")
    side_frame = Frame(root,height=50, bg="navy")
    side_frame.pack(side="top", fill="x")
    side_frame = Frame(root, height=50, bg="navy")
    side_frame.pack(side="bottom", fill="x")
    side_frame = Frame(root, width=50, bg="navy")
    logo_label = Label(root, image=logo_photo)
    logo_label.pack() 
    
    student_button = Button(root, text="Show Student Data", command=show_student_data, bd=10,highlightbackground="red")
    student_button.place(relx=0.5, rely=0.4, anchor="center")

    holiday_button = Button(root, text="Show Upcoming Holidays", command=show_upcoming_holidays,bd=10,highlightbackground="red")
    holiday_button.place(relx=0.5, rely=0.5, anchor="center")

    event_button = Button(root, text="Show Upcoming Events", command=show_upcoming_events, bd=10,highlightbackground="red")
    event_button.place(relx=0.5, rely=0.6, anchor="center")

    exam_button = Button(root, text="Show Exam Schedule", command=show_exam_schedule, bd=10,highlightbackground="red")
    exam_button.place(relx=0.5, rely=0.7, anchor="center")

    root.mainloop()

def sendEmail_terminal(to, sub, msg):
    print(f"\n Email to {to} sent with subject: {sub} and message: \n {msg}")  ## here we are checking whether email is getting sent in email

def sendEmail_SMTP(to, sub, msg): 

    smtp_conn = smtplib.SMTP('smtp-mail.outlook.com', 587)

    status_code, response = smtp_conn.ehlo()
    print(f"[*] Echoing the server: {status_code}, {response}")

    status_code, response = smtp_conn.starttls()
    print(f"[*] Starting TLS connection: {status_code}, {response}")

    status_code, response = smtp_conn.login(EMAIL_ID,EMAIL_PSWD)
    print(f"[*] Logging in: {status_code}, {response}")

    smtp_conn.sendmail(EMAIL_ID, to, f"Subject: {sub} \n\n {msg}")
    smtp_conn.quit
    print("mail sent succesful")


if __name__ == "__main__":

    df= createExcel()
    df= pd.read_csv('StudList.csv')
    # print(df)

    today= dt.datetime.now().strftime("%d-%m-%y") 
    # print("today's date: ",today)

    hl= holidayList()
    hl=  pd.read_csv('HolidayList.csv')
    # print("--------Holiday List---------")
    # print(hl)
    
    EL= EventList()
    EL=  pd.read_csv('EventList.csv')
    # print("--------EVent List---------")
    # print(EL)

    ES= ExamSchedule()
    ES=  pd.read_csv('ExamSchedule.csv')
    # print("--------Exam Schedule---------")
    # print(ES)

    # mid term examination news
    if(today=='26-03-24'):
        msg1='''Hello student, this is to announce that your mid term will be schedule on 1st april,2024. 
        time table pdf will be share soon. \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Mid Term Announcement', msg1)
            sendEmail_SMTP(item['Email'],' Vacation', msg1)
            
    # end term announcement
    if(today=='06-04-24'):
        msg2 ='''Hello student, This message is to announce you that you end term will be starting from 15th may,2024.
        Time table will be shared soon. \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],'End Semester Announcement', msg2)
            sendEmail_SMTP(item['Email'],' Vacation', msg2)

    # mini project deadline alerts
    if(today=='12-04-24'):
        msg3 ='''Hello student, This message is to announce you that all mini project deadline will be on 25th April,2024
         please complete your submission on time as final evaluation will be done on 25th april.  \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Mid Term Announcement', msg3)
            sendEmail_SMTP(item['Email'],' Vacation', msg3)

    # practical schedule
    if(today=='10-04-24'):
        msg4 ='''Hello student, This message is to announce you that you end term practical exam will be starting from 26th April,2024. 
        Time table will be shared soon. \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Practical Examination Announcement', msg4)
            sendEmail_SMTP(item['Email'],' Vacation', msg4)

    # SIH announcement
    if(today=='18-08-24'):
        msg5 ='''Smart India Hackathon will be conducting in MIT-WPU university on the date of 10th August 2024.
        \n Venue: tikona Hall \n: date: 15th september,2024 \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Smart India Hackathon', msg5)
            sendEmail_SMTP(item['Email'],' Vacation', msg5)

    # year end vacaion 
    if(today=='01-05-24'):
        msg6 ='''Year end vacation will be starting from 1st June,2024 to 31st July, 2024.
          New acedemy session will be starting from !st August 2024. 
          !st installment of fees of next year will begin from 1st July,2024. Remaining update will be shared soon \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Vacation', msg6)
            sendEmail_SMTP(item['Email'],' Vacation', msg6)

   # ride event announement
    if(today=='31-08-24'):
        msg7 ='''RIDE-EVENT will be starting from 1st October, 2024.
        you will be seeing many entrepreneu, different events like Pro-nite, Youtube fest, Comedy-nights
         special guest like gaurav kapoor, Namita Thapar, BhuvamBam are invited.
         hope to see you all \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Smart India Hackathon', msg7)
            sendEmail_SMTP(item['Email'],' Vacation', msg7)

    # semester end news
    if(today=='12-12-24'):
        msg8 ='''Semester end vacation will be starting from 29th December,2024 to 16th January, 2024.
          New Semester session will be starting from 17th January,2024  \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Vacation', msg8)
            sendEmail_SMTP(item['Email'],' Vacation', msg8)

    # independence holiday
    if(today=='09-12-24'):
        msg9 ='''This is to announce you that, Huge Independence function will be taken in MIT-ADT.
          Various of university are joining there and participating in event. all university student have 
            to perform one group function Either group dance, Concert Band, March Past related to patriatism.
              Intrested one please Fill the google form that will be shared soon.
                Hope you all will attend the function and enjoy it.  \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Vacation', msg9)
            sendEmail_SMTP(item['Email'],' Vacation', msg9)

    # placement and internships round
    if(today=='24-04-24'):
        msg10= ''' All various companies are coming in campus for placement(Round-1) and Summer internship from 1st June onwards.
        intrested people please appl form given in tcs ion.  \n Thank you'''
        for index,item in df.iterrows():
            print(f"\n\n\n {index}")
            sendEmail_terminal(item['Email'],' Vacation', msg10)
            sendEmail_SMTP(item['Email'],' Vacation', msg10)
    
  #  sending a mail to defaulters
    if(today=='04-05-24'):
        msg_defaulters= ''' This mail is rearding to the defaulter students, Those who have less than 75 percent
                   Have to meet me in my Cabin Dr 007 \n Thank You'''
        for index,item in df.iterrows():
            if item['Attendence'] <= 75:
               print(f"\n\n\n {index}")
               sendEmail_terminal(item['Email'],' Vacation', msg_defaulters)
               sendEmail_SMTP(item['Email'],' Vacation', msg_defaulters)

   # Scholarships
    if(today=='18-12-24'):
        msg_ScholarShip1= ''' Congratulation you have got Scholarship of 75% As you have scored more than 9.0 Cgpa
                             Please Visit department office and fill the form. In case you have paid the university fees,
                             You will be refunded on the basis of form. \n Thank You'''
        msg_ScholarShip2= ''' Congratulation you have got Scholarship of 50% As you have scored more than 8.0 Cgpa
                             Please Visit department office and fill the form. In case you have paid the university fees,
                             You will be refunded on the basis of form. \n Thank You'''
        for index,item in df.iterrows():
            if item['CGPA'] >= 9.0:
               print(f"\n\n\n {index}")
               sendEmail_terminal(item['Email'],' Vacation', msg_ScholarShip1)
               sendEmail_SMTP(item['Email'],' Vacation', msg_ScholarShip1)
            if item['CGPA'] >= 8.0:
               print(f"\n\n\n {index}")
               sendEmail_terminal(item['Email'],' Vacation', msg_ScholarShip2)
               sendEmail_SMTP(item['Email'],' Vacation', msg_ScholarShip2)

    print("----------------ADDITIONAL FEATURES-------------------")
    print("----------------TO DISPLAY A WEBSITE------------------")
    otherdetails=input("Enter ok to display(else no to exit): ")
    if(otherdetails=='ok'):
        display_data()

