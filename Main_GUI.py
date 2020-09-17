from tkinter import *
from PIL import ImageTk,Image
import pandas as pd 
import tkinter.messagebox as msg
import os
import xlsxwriter


global login_by
login_by = None 

global student_details
student_details = []

root = Tk()
root.geometry("580x500")
root.maxsize(580,550)
root.minsize(580,550)
root.title("Auto Attendence System")
root.iconbitmap('G:/Projects/Attendence System/icon.ico')
image = Image.open("G:/Projects/Attendence System/Background.jpg")
image = image.resize((580,600), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(image)


log_in = Label(image=photo)

#user for login

global Users

#####These are the Users########
Users = {"ajit" : "153",
         "arvind" : "243"
        }


##### Code for Authentication #########
def login():
    name = lid.get()
    code = lpass.get()
    for username,password in Users.items():
        if name==username and code==password:
            login_by = username
            login_allow()
            break
        else:
            login_allow()
            '''message = msg.showwarning("This is a Warning","Enter valid Details")
            Label(root,text=message).pack()
            break'''
        


########### Code for the Login Page Start here #################
f_login = Frame(log_in,padx=25,pady=25)
lb0 = Label(f_login,text="Enter Details",bg="orange",fg="blue",font="lucide 10 bold",width=35,pady=4).grid(columnspan=3,row=0,pady =15)
lb1=Label(f_login,text="Enter ID: ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
lid=StringVar()
e1=Entry(f_login,textvariable=lid,width="28").grid(column=1,row=2)
lb2=Label(f_login,text="Enter Password: ",font="lucida 10 bold").grid(column=0,row=3,pady="4")
lpass=StringVar()
e2=Entry(f_login,textvariable=lpass,width="28").grid(column=1,row=3)
btn=Button(f_login,text="login",bg="green",fg="white",width="10",font="lucida 10 bold",command=login)
btn.grid(columnspan=3,row=5,pady="10")
f_login.pack(pady="165")
############### Code for the Login Page ends here ##############
log_in.pack(ipadx="100",fill=BOTH)


###code for more #####
def more(z): 
     if z==1:
        l.pack_forget()
        l1.pack(ipadx="100",fill = BOTH)
     elif z==2:
         message = msg.showinfo("Auto Attendence System","Auto Attendence System \n Made by Ajit Verma")
         Label(root,text=message).pack()
         if message=="ok":
             message.destroy()

def student(x):
    l1.pack_forget()
    l.pack(ipadx="120",fill=BOTH)
    if (x==1):
        f4.pack_forget()
        f2.pack_forget()
        f21.pack_forget()
        f3.pack_forget()
        f31.pack_forget()
        f1.pack(pady="100")
    if (x==2):
        f4.pack_forget()
        f1.pack_forget()
        f21.pack_forget()
        f3.pack_forget()
        f31.pack_forget()
        f2.pack(pady="120")
    if (x==3):
        f1.pack_forget()
        f2.pack_forget()
        f21.pack_forget()
        f4.pack_forget()
        f31.pack_forget()
        f3.pack(pady="120")
    if (x==4):
        f4.pack(pady="120")
        f1.pack_forget()
        f3.pack_forget()
        f31.pack_forget()
        f2.pack_forget()
        f21.pack_forget()
        
    


def clear():
    response = msg.showinfo("Auto Attendence System","Details are successfully Registered")
    #Label(root,text=response).pack()
    # response=="ok":
        #response.destroy()
    name.delete(0,END)
    enroll.delete(0,END)
    course.delete(0,END)
    section.delete(0,END)
    semester.delete(0,END)
    contact.delete(0,END)
    email.delete(0,END)

def excel():
    file = pd.read_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx")
    writer = pd.ExcelWriter('G:/Projects/Attendence System/Excel Files/Student_details.xlsx', engine='xlsxwriter')
    new = file.append(df,ignore_index=True)
    new = new.sort_values('Rollno')
    new.to_excel(writer, index=False, sheet_name='Section A')
    workbook = writer.book
    worksheet = writer.sheets['Section A']
    worksheet.set_zoom(100)
    frmt = workbook.add_format({'align':'center'})
    worksheet.set_column('A:A', 20,frmt)
    worksheet.set_column('B:B', 12,frmt)
    worksheet.set_column('C:E', 10,frmt)
    worksheet.set_column('F:F', 18,frmt)
    worksheet.set_column('G:G',30,frmt)
    worksheet.freeze_panes(1,0)
    writer.save()
    #os.startfile('G:/Projects/Attendence System/Excel Files/Student_details.xlsx')
       


def Register():
    naam = name.get()
    roll_no = enroll.get()
    coe = course.get()
    sec = section.get()
    seme = semester.get()
    con = contact.get()
    em = email.get()

    """if naam and roll_no and coe and sec and seme and con and em is NULL:
        response = msg.showwarning("Auto Attendence System","Please Fill all the Entries")
        Label(root,text=response)"""

    l = [naam,roll_no,coe,sec,seme,con,em]
    columns = ["Name","Rollno","Course","Section","Semester","Contact no","Email id"]
    student_details.append(l)

    global df
    df = pd.DataFrame(student_details,columns = columns)
    #df.to_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx",index=False)
    excel()
    clear()

#code for the login purpose
def login_allow():
    log_in.pack_forget()

#############Code for Menu Start Here #####################################
    mainmenu = Menu(root)
    
    m1 = Menu(mainmenu, tearoff=0)
    m1.add_command(label="Register", command=lambda:student(1))
    m1.add_command(label="View", command=lambda:student(2))
    m1.add_separator()
    m1.add_command(label="Update", command=lambda:student(3))
    m1.add_command(label="Delete", command=lambda:student(4))
    mainmenu.add_cascade(label="Student", menu=m1)

    m2 = Menu(mainmenu, tearoff=0)
    m2.add_command(label="Detect", command=None)
    m2.add_separator()
    m2.add_command(label="View Excel", command=None)
    mainmenu.add_cascade(label="Attendance", menu=m2)

    

    m3 = Menu(mainmenu, tearoff=0)
    m3.add_command(label="Help", command=lambda:more(1))
    m3.add_command(label="About Us", command=lambda: more(2))
    mainmenu.add_cascade(label="More", menu=m3)
   
    root.config(menu=mainmenu)

 #######Code for Menu ends Here #############   


########## code for the registration starts here ###############
l = Label(image = photo)

f1=Frame(l,pady="5",padx="25")
l0=Label(f1,text="Registration Form",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")
l1=Label(f1,text="Name : ",font="lucida 10 bold").grid(column=0,row=1,pady="4")
name=Entry(f1,width="28")
name.grid(column=1,row=1)
l2=Label(f1,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
enroll=Entry(f1,width="28")
enroll.grid(column=1,row=2)
l3=Label(f1,text="Course : ",font="lucida 10 bold").grid(column=0,row=3,pady="4")
course=Entry(f1,width="28")
course.grid(column=1,row=3)

l32=Label(f1,text="Section : ",font="lucida 10 bold").grid(column=0,row=4,pady="4")
section=Entry(f1,width="28")
section.grid(column=1,row=4)

l33=Label(f1,text="Sem : ",font="lucida 10 bold").grid(column=0,row=5,pady="4")
semester=Entry(f1,width="28")
semester.grid(column=1,row=5)

l5=Label(f1,text="Contact No : ",font="lucida 10 bold").grid(column=0,row=6,pady="4")
contact=Entry(f1,width="28")
contact.grid(column=1,row=6)

l6=Label(f1,text="Email : ",font="lucida 10 bold").grid(column=0,row=7,pady="4")
email=Entry(f1,width="28")
email.grid(column=1,row=7)
btn=Button(f1,text="Submit",bg="green",fg="white",width="10",font="lucida 10 bold",command=Register)
btn.grid(columnspan=3,row=8,pady="10")
f1.pack(pady="100")

#################### code for Registration ends here ############

def back():
    f21.pack_forget()
    f2.pack(pady="120")
    


def view():
    f2.pack_forget()
    df2 = pd.read_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx",index=False)
    data = df2.loc[df2["Rollno"]== venroll.get()]
    n1 = data["Name"]
    r1 = data["Rollno"]
    co1 = data["Course"]
    sec1 = data["Section"]
    sem1 = data["Semester"]
    con1 = data["Contact no"]
    em1 = data["Email id"]
    dname.set(*n1)
    dRoll_no.set(*r1)
    dcourse.set(*co1)
    dsection.set(*sec1)
    dsemester.set(*sem1)
    dcontact.set(*con1)
    demail.set(*em1)
    f21.pack(pady="90")
    

#################### code for student details start here #########
f2=Frame(l,pady="25",padx="25")
l0=Label(f2,text="Student details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")
l2=Label(f2,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
venroll = StringVar()
e2=Entry(f2,textvariable=venroll,width="28").grid(column=1,row=2)
l4=Label(f2,text="Semester",font="lucida 10 bold").grid(column=0,row=4,pady="4")
vsem = StringVar()
vsem.set("1st sem") # default value
w1 = OptionMenu(f2,vsem,"1st sem","2nd sem","3rd sem","4th sem","5th sem","6th sem","7th sem","8th sem").grid(column=1,row=4,pady="4")
l5=Label(f2,text="Section",font="lucida 10 bold").grid(column=0,row=5,pady="4")
vsection = StringVar()
vsection.set("CSE-A") # default value
w2 = OptionMenu(f2,vsection, "CSE-A", "CSE-B").grid(column=1,row=5,pady="4")
btn=Button(f2,text="OK",bg="green",fg="white",width="10",font="lucida 10 bold",command=view)
btn.grid(columnspan=3,row=7,pady="20")
f2.pack(pady="115")
#################### Code for Student details ends here #######################


######################## Student detail view ##################################
f21=Frame(l,pady="20",padx="25")   
l0=Label(f21,text="Student Details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")   
l1=Label(f21,text="Name : ",font="lucida 10 bold").grid(column=0,row=1,pady="4")
dname = StringVar()
l11=Label(f21,textvariable=dname,width="28").grid(column=1,row=1)
l2=Label(f21,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
dRoll_no = StringVar()
l22=Label(f21,textvariable=dRoll_no,width="28").grid(column=1,row=2)
l3=Label(f21,text="Course : ",font="lucida 10 bold").grid(column=0,row=3,pady="4")
dcourse = StringVar()
l33=Label(f21,textvariable=dcourse,width="28").grid(column=1,row=3)
l4=Label(f21,text="Section : ",font="lucida 10 bold").grid(column=0,row=4,pady="4")
dsection = StringVar() 
l44=Label(f21,textvariable=dsection,width="28").grid(column=1,row=4)
l5=Label(f21,text="Semseter : ",font="lucida 10 bold").grid(column=0,row=5,pady="4")
dsemester = StringVar()
l55=Label(f21,textvariable=dsemester,width="28").grid(column=1,row=5)
l6=Label(f21,text="Contact No : ",font="lucida 10 bold").grid(column=0,row=6,pady="4")
dcontact = StringVar()
l66=Label(f21,textvariable=dcontact,width="28").grid(column=1,row=6)
l7=Label(f21,text="Email : ",font="lucida 10 bold").grid(column=0,row=7,pady="4")
demail = StringVar()
l77=Label(f21,textvariable=demail,width="28").grid(column=1,row=7)
btn=Button(f21,text="Back",bg="green",fg="white",width="10",font="lucida 10 bold",command=back)
btn.grid(columnspan=3,row=9,pady="8")
f21.pack(pady="100")   
############################################################

''' code for update student details  is start here     '''
def update():
    f3.pack_forget()
    file = pd.read_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx")
    data = file.loc[file["Rollno"]==urollno.get()]
    n2 = data["Name"]
    r2 = data["Rollno"]
    co2 = data["Course"]
    sec2 = data["Section"]
    sem2 = data["Semester"]
    con2 = data["Contact no"]
    em2 = data["Email id"]
    f31.pack(pady="80")
    e1.insert(0,*n2)
    e2.insert(0,*r2)
    e3.insert(0,*co2)
    e4.insert(0,*sec2)
    e5.insert(0,*sem2)
    e6.insert(0,*con2)
    e7.insert(0,*em2)
    

def cancel():
    f31.pack_forget()
    f3.pack(pady="120")
def update_details():
    file = pd.read_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx")
    file.loc[file["Rollno"]==urollno.get(),"Name"] = uname.get()
    file.loc[file["Rollno"]==urollno.get(),"Rollno"] = uname.get()
    file.loc[file["Rollno"]==urollno.get(),"Course"] = ucourse.get()
    file.loc[file["Rollno"]==urollno.get(),"Section"] = usection.get()
    file.loc[file["Rollno"]==urollno.get(),"Semester"] = usemester.get()
    file.loc[file["Rollno"]==urollno.get(),"Contact no"] = ucontact.get()
    file.loc[file["Rollno"]==urollno.get(),"Email id"] = uemail.get()

    writer = pd.ExcelWriter('G:/Projects/Attendence System/Excel Files/Student_details.xlsx', engine='xlsxwriter')
    file = file.sort_values('Rollno')
    file.to_excel(writer, index=False, sheet_name='Section A')
    workbook = writer.book
    worksheet = writer.sheets['Section A']
    worksheet.set_zoom(100)
    frmt = workbook.add_format({'align':'center'})
    worksheet.set_column('A:A', 20,frmt)
    worksheet.set_column('B:B', 12,frmt)
    worksheet.set_column('C:E', 10,frmt)
    worksheet.set_column('F:F', 18,frmt)
    worksheet.set_column('G:G',30,frmt)
    worksheet.freeze_panes(1,0)
    writer.save()

    print("details updated")
    

f31=Frame(l,pady="25",padx="25")
l0=Label(f31,text="Update Details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15") 
l1=Label(f31,text="Name : ",font="lucida 10 bold").grid(column=0,row=1,pady="4")

uname=StringVar()
e1=Entry(f31,textvariable=uname,width="28")
e1.grid(column=1,row=1) 
l2=Label(f31,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")

uenrollment=StringVar()
e2=Entry(f31,textvariable=uenrollment,width="28")
e2.grid(column=1,row=2) 
l3=Label(f31,text="Course : ",font="lucida 10 bold").grid(column=0,row=3,pady="4")

ucourse=StringVar()   
e3=Entry(f31,textvariable=ucourse,width="28")
e3.grid(column=1,row=3) 
l4=Label(f31,text="Section : ",font="lucida 10 bold").grid(column=0,row=4,pady="4")

usection=StringVar()   
e4=Entry(f31,textvariable=usection,width="28")
e4.grid(column=1,row=4) 
l5=Label(f31,text="Semester : ",font="lucida 10 bold").grid(column=0,row=5,pady="4")

usemester=StringVar()
e5=Entry(f31,textvariable=usemester,width="28")
e5.grid(column=1,row=5) 
l6=Label(f31,text="Contact No : ",font="lucida 10 bold").grid(column=0,row=6,pady="4")

ucontact=StringVar()
e6=Entry(f31,textvariable=ucontact,width="28")
e6.grid(column=1,row=6) 
l7=Label(f31,text="Email : ",font="lucida 10 bold").grid(column=0,row=7,pady="4")

uemail=StringVar()
e7=Entry(f31,textvariable=uemail,width="28")
e7.grid(column=1,row=7) 

btn=Button(f31,text="Submit",bg="green",fg="white",width="10",font="lucida 10 bold",command=update_details)
btn.grid(column=0,row=8,pady="20")    
btn1=Button(f31,text="cancel",bg="green",fg="white",width="10",font="lucida 10 bold",command=cancel)
btn1.grid(column=1,row=8,pady="20")    
f31.pack(pady="100")


f3=Frame(l,pady="25",padx="25")
l0=Label(f3,text="Update details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")
l2=Label(f3,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
urollno = StringVar()
e2=Entry(f3,textvariable= urollno,width="28").grid(column=1,row=2)

"""l3=Label(f3,text="Year",font="lucida 10 bold").grid(column=0,row=3,pady="4")
year = StringVar()
year.set("1st Year") # default value
w = OptionMenu(f3,year, "1st Year", "2nd Year", "3rd Year","4th Year").grid(column=1,row=3,pady="4")"""

l4=Label(f3,text="Sem",font="lucida 10 bold").grid(column=0,row=4,pady="4")
sem = StringVar()
sem.set("1st sem") # default value
w1 = OptionMenu(f3,sem,"1st sem","2nd sem","3rd sem","4th sem","5th sem","6th sem","7th sem","8th sem").grid(column=1,row=4,pady="4")

l5=Label(f3,text="Section",font="lucida 10 bold").grid(column=0,row=5,pady="4")
section = StringVar()
section.set("CSE-A") # default value
w2 = OptionMenu(f3,section, "CSE-A", "CSE-B").grid(column=1,row=5,pady="4")

btn=Button(f3,text="OK",bg="green",fg="white",width="10",font="lucida 10 bold",command=update)
btn.grid(columnspan=3,row=7,pady="20")
f3.pack(pady="115")

''' code for update student details  is end here     '''



'''delete Student detail'''
def delete():
    value = msg.askquestion("Delete student delails", "Are you sure")
    if value == "yes":
        file = pd.read_excel("G:/Projects/Attendence System/Excel Files/Student_details.xlsx")
        writer = pd.ExcelWriter('G:/Projects/Attendence System/Excel Files/Student_details.xlsx', engine='xlsxwriter')
        file = file[file.Rollno!=denroll.get()]
        file.to_excel(writer, index=False, sheet_name='Section A')
        workbook = writer.book
        worksheet = writer.sheets['Section A']
        worksheet.set_zoom(100)
        frmt = workbook.add_format({'align':'center'})
        worksheet.set_column('A:A', 20,frmt)
        worksheet.set_column('B:B', 12,frmt)
        worksheet.set_column('C:E', 10,frmt)
        worksheet.set_column('F:F', 18,frmt)
        worksheet.set_column('G:G',30,frmt)
        worksheet.freeze_panes(1,0)
        writer.save()


f4=Frame(l,pady="25",padx="25")
l0=Label(f4,text="Delete details",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(columnspan=3,row=0,pady="15")
l2=Label(f4,text="Roll No : ",font="lucida 10 bold").grid(column=0,row=2,pady="4")
denroll=StringVar()
denl=Entry(f4,textvariable=denroll,width="28").grid(column=1,row=2)

#l3=Label(f4,text="Year",font="lucida 10 bold").grid(column=0,row=3,pady="4")
#dyear = StringVar()
#dyear.set("1st Year") # default value
#w = OptionMenu(f4,dyear, "1st Year", "2nd Year", "3rd Year","4th Year").grid(column=1,row=3,pady="4")

l4=Label(f4,text="Sem",font="lucida 10 bold").grid(column=0,row=4,pady="4")
dsem = StringVar()
dsem.set("1st sem") # default value
w1 = OptionMenu(f4,dsem,"1st sem","2nd sem","3rd sem","4th sem","5th sem","6th sem","7th sem","8th sem").grid(column=1,row=4,pady="4")

l5=Label(f4,text="Section",font="lucida 10 bold").grid(column=0,row=5,pady="4")
dsection = StringVar()
dsection.set("CSE-A") # default value
w2 = OptionMenu(f4,dsection, "CSE-A", "CSE-B").grid(column=1,row=5,pady="4")

btn=Button(f4,text="Delete",bg="green",fg="white",width="10",font="lucida 10 bold",command=delete)
btn.grid(columnspan=3,row=7,pady="20")
f4.pack(pady="115")
###############################################
l.pack(ipadx="110",fill=BOTH)

l1  = Label(image=photo)
f=Frame(l1,pady="25",padx="25")
lbl=Label(f,text="we will help you",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(column=0,row=0)
lbl=Label(f,text="you can contact us on following",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(column=0,row=1)
lbl=Label(f,text="Email  :  ajitverma1503@gmail.com",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(column=0,row=2)
lbl=Label(f,text="mobile : 1800-6512-154",bg="orange",fg="blue",font="lucida 10 bold",width="35",pady="4").grid(column=0,row=3)
f.pack(pady="185")

l1.pack(ipadx="100",fill=BOTH)

root.mainloop()

