import smtplib
import threading
import tkinter as tk
from tkinter.ttk import *
from tkinter import filedialog
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication



def gui():
    # email sender information
    from_address = ''
    password = ''
    smtp_server = ''
    smtp_port=0
    # Email Content
    email_list =[]
    #list_filepath=""
    attachment_filename=""
    email_subject = ""
    email_body=""
    # process
    email_sent=0
    Total_emails=0
    i=1

    # window setting
    window = tk.Tk()
    window.geometry("540x460")
    icon = tk.PhotoImage(file="mail.png")
    window.iconphoto(True, icon)
    window.title("Bulk Email Sender")
    # Message Box
    tk.messagebox.showinfo(title='How To',message=
    """    1.Enter Email,Password,smpt server,smtp port
    2.Load text file containing email list Note: .txt only
    3.Load Attachment Note: .png only
    4.Edit Email Subject and edit body
    5.Click Send and wait for completion""")
    # Labels

    email_label= tk.Label(window,text="Email")# email
    email_label.place(x=10,y=10)
    pass_label= tk.Label(window,text="Password")# password
    pass_label.place(x=280,y=10)
    server_label= tk.Label(window,text="SMTP Server")# smtp server
    server_label.place(x=10,y=40)
    port_label= tk.Label(window,text="SMTP Port")# smtp port
    port_label.place(x=280,y=40)
    email_list_label= tk.Label(window,text="Load Email List (.txt)",font=("","10", "bold"))# txt path
    email_list_label.place(x=10,y=85)
    list_load_label= tk.Label(window,text="Not Loaded",foreground="black",font=("","10", "bold"))# load prompt
    list_load_label.place(x=300,y=85)
    attach_label= tk.Label(window,text="Load Attachment (.png)",foreground="black",font=("","10", "bold"))# attachment
    attach_label.place(x=10,y=140)
    attach_status_label= tk.Label(window,text="Not Loaded",foreground="black",font=("","10", "bold"))# attachment load prompt
    attach_status_label.place(x=350,y=140)
    email_subject_label=tk.Label(window,text="Email Subject",foreground="black",font=("","10", "bold"))# subject
    email_subject_label.place(x=10,y=180)
    email_body_label=tk.Label(window,text="Email Body",foreground="black",font=("","10", "bold"))# email body
    email_body_label.place(x=10,y=220)
    counter_label=tk.Label(window,text="",foreground="black",font=("","10", "bold"))# counter
    counter_label.place(x=20,y=270)

    #Entry Boxes
    email_entry = tk.Entry(window,bd=4,width=25,font=("Times","9", "bold") )# email entry
    email_entry.place(x=100,y=10)
    pass_entry = tk.Entry(window,bd=4,width=25,font=("Times","9", "bold"),show="*")# password entry
    pass_entry.place(x=360,y=10)
    server_entry = tk.Entry(window,bd=4,width=25,font=("Times","9", "bold"))# server entry
    server_entry.place(x=100,y=40)
    port_entry = tk.Entry(window,bd=4,width=25,font=("Times","9", "bold"))# port entry
    port_entry.place(x=360,y=40)
    subject_entry = tk.Entry(window,bd=4,width=40,font=("Times","9", "bold"))# subject entry
    subject_entry.place(x=120,y=180)

    def btn_click():
        nonlocal attachment_filename,email_body,email_list
        incr = (1/Total_emails)*100
        # getting  data from fields
        from_address=email_entry.get()
        password=pass_entry.get()
        smtp_server=server_entry.get()
        smtp_port=int(port_entry.get())
        email_subject=subject_entry.get()

        # create email message
        msg = MIMEMultipart()
        msg['From'] = from_address
        msg['Subject'] = email_subject
        msg.attach(MIMEText(email_body, 'plain'))
        # attach file
        with open(attachment_filename, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='png')
            attachment.add_header('content-disposition', 'attachment', filename=attachment_filename)
            msg.attach(attachment)

        try:
            # create smtp connection
            server = smtplib.SMTP_SSL(smtp_server, smtp_port)
            #server.starttls()
            server.ehlo()
            server.login(from_address, password)
            print("logged in")
            log("logged in .....")

            # send email to each recipient
            for email in email_list:
                try:
                    # create a new copy of message for each recipient
                    msg_to_send = MIMEMultipart()
                    msg_to_send.attach(msg)
                    msg_to_send['To'] = email

                    # send email
                    server.sendmail(from_address, email, msg_to_send.as_string())
                    window.update_idletasks()
                    print(f'Email sent to {email}')
                    log(f'Email sent to {email}')
                    nonlocal email_sent
                    email_sent+=1
                    bar["value"]+=incr
                    counter_label["text"]=str(email_sent)+'/'+str(Total_emails)
                except Exception as e:
                    print(f'Error sending email to {email}: {e}')
                    log(f'Error sending email to {email}: {e}')

            # close smtp connection
            server.quit()
            counter_label["text"]+=" Completed !"
            log("Email Sent To All Addresses")
        except Exception as e:
            print("Error Logging in")
            log("Error Logging in....")

    def load_file():
        options = {
        'defaultextension': '.txt',
        'filetypes': [('Text Files', '*.txt')],}
        email_file=filedialog.askopenfile(mode='r', **options)
        if(email_file):
            list_load_label["text"]="Loaded"
            list_load_label["foreground"]="green"
            nonlocal email_list
            email_list = [line.strip() for line in email_file.readlines()]
            nonlocal Total_emails
            Total_emails += len(email_list)
            counter_label["text"]=str(email_sent)+'/'+str(Total_emails)
            print(email_list)
            log("Email List Loaded")
            email_file.close()
    def load_attachment():
        nonlocal attachment_filename
        options = {
        'defaultextension': '.png',
        'filetypes': [('png files', '*.png')],}
        attachment_filename=filedialog.askopenfilename(**options)
        if(attachment_filename!=""):
            attach_status_label["text"]="Loaded"
            attach_status_label["foreground"]="green"
            print(attachment_filename)
            log("Attached file")
    def edit_email_body():
        email_body_window=tk.Toplevel();
        # Edit Window
        email_body_window.geometry("480x580")
        email_body_window.title("Edit Email Body")
        # Labels
        Heading=tk.Label(email_body_window,text="Email Body",foreground="black",font=("","15", "bold"))# subject
        Heading.place(x=10,y=10)
        # Text area
        text = tk.Text(email_body_window,width=57,height=30)
        text.place(x=10,y=55)
        # Button command
        def ok_click():
            nonlocal email_body
            email_body = text.get("1.0",tk.END)
            email_body_window.destroy()
        # Button
        ok_btn=tk.Button(email_body_window,command=ok_click,text="OK",font=("","10",""),bd=5,bg="#bfc6c4",activebackground="black",activeforeground="white")
        ok_btn.pack(side=tk.BOTTOM,pady=5,padx=20,fill="x")
    # helper functions
    def log(text):
        nonlocal i
        console.insert(i,text)
        i+=1

    # Button Send
    send_button= tk.Button(window,
                            command=threading.Thread(target=btn_click).start,
                            text="Send",font=("Times","15","italic"),
                            bd=5,bg="#bfc6c4",activebackground="black",
                            activeforeground="white")

    send_button.pack(side=tk.BOTTOM,pady=5,padx=20,fill="x")

    # Button Load File
    load_button= tk.Button(window,command=load_file,text="Load File",bd=5,bg="#bfc6c4",activebackground="black",activeforeground="white")
    load_button.place(x=150,y=80,width=100)
    # Button Load Attachment File
    load_button= tk.Button(window,command=load_attachment,text="Load Attachment File",bd=5,bg="#bfc6c4",activebackground="black",activeforeground="white")
    load_button.place(x=170,y=135,width=140)
    # Button Edit email body
    email_body_button= tk.Button(window,command=edit_email_body,text="Edit Email Body",bd=5,bg="#bfc6c4",activebackground="black",activeforeground="white")
    email_body_button.place(x=170,y=210,width=140)
    # List box

    # Progress Bar
    bar =Progressbar(window,orient="horizontal",length=500)
    bar.place(x=20,y=250)
    #List Box

    console = tk.Listbox(window,height=7,font=("Arial","11",""),width=85,background="black",foreground="#39FF14")
    console.place(x=10,y=290)
    # Main Loop
    window.mainloop()



gui()
