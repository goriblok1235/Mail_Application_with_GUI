from tkinter import *
from email.message import EmailMessage
from tkinter import messagebox, filedialog
import smtplib
from pygame import mixer
import os
import imghdr
import pandas


check = False


def browse():
    global final_emails
    path = filedialog.askopenfilename(
        initialdir='c:/', title='Select Excel File')
    if path == '':
        messagebox.showerror('Error', 'Please select an Excel File')

    else:
        data = pandas.read_excel(path)
        if 'Email' in data.columns:
            emails = list(data['Email'])
            final_emails = []
            for i in emails:
                if pandas.isnull(i) == False:
                    final_emails.append(i)

            if len(final_emails) == 0:
                messagebox.showerror(
                    'Error', 'File does not contain any email addresses')

            else:
                toEntryField.config(state=NORMAL)
                toEntryField.insert(0, os.path.basename(path))
                toEntryField.config(state='readonly')
                total_Label.config(text='Total: '+str(len(final_emails)))
                sent_Label.config(text='Sent:')
                left_Label.config(text='Left:')
                failed_Label.config(text='Failed:')


def button_check():
    if choice.get() == 'multiple':
        browseButton.config(state=NORMAL)
        toEntryField.config(state='readonly')

    if choice.get() == 'single':
        browseButton.config(state=DISABLED)
        toEntryField.config(state=NORMAL)


def attachment():
    global filename, filetype, filepath, check
    check = True

    filepath = filedialog.askopenfilename(
        initialdir='c:/', title='Select File')
    filetype = filepath.split('.')
    filetype = filetype[1]
    filename = os.path.basename(filepath)
    text_area.insert(END, f'\n{filename}\n')


def sendingEmail(toAddress, subject, body):
    f = open('settings.txt', 'r')
    for i in f:
        credentials = i.split(',')

    message = EmailMessage()
    message['subject'] = subject
    message['to'] = toAddress
    message['from'] = credentials[0]
    message.set_content(body)
    if check:
        if filetype == 'png' or filetype == 'jpg' or filetype == 'jpeg':
            f = open(filepath, 'rb')
            file_data = f.read()
            subtype = imghdr.what(filepath)

            message.add_attachment(
                file_data, maintype='image', subtype=subtype, filename=filename)

        else:
            f = open(filepath, 'rb')
            file_data = f.read()
            message.add_attachment(
                file_data, maintype='application', subtype='octet-stream', filename=filename)

    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(credentials[0], credentials[1])
    s.send_message(message)
    x = s.ehlo()
    if x[0] == 250:
        return 'sent'
    else:
        return 'failed'


def send_email():
    if toEntryField.get() == '' or subjectEntry.get() == '' or text_area.get(1.0, END) == '\n':
        messagebox.showerror('Error', 'All Fields Are Required', parent=root)

    else:
        if choice.get() == 'single':
            result = sendingEmail(
                toEntryField.get(), subjectEntry.get(), text_area.get(1.0, END))
            if result == 'sent':
                messagebox.showinfo('Success', 'Email is sent successfulyy')

            if result == 'failed':
                messagebox.showerror('Error', 'Email is not sent.')

        if choice.get() == 'multiple':
            sent = 0
            failed = 0
            for x in final_emails:
                result = sendingEmail(
                    x, subjectEntry.get(), text_area.get(1.0, END))
                if result == 'sent':
                    sent += 1
                if result == 'failed':
                    failed += 1

                total_Label.config(text='')
                sent_Label.config(text='Sent:' + str(sent))
                left_Label.config(
                    text='Left:' + str(len(final_emails) - (sent + failed)))
                failed_Label.config(text='Failed:' + str(failed))

                total_Label.update()
                sent_Label.update()
                left_Label.update()
                failed_Label.update()

            messagebox.showinfo('Success', 'Emails are sent successfully')


def settings():
    def clear1():
        fromEntryField.delete(0, END)
        passwordEntryField.delete(0, END)

    def save():
        if fromEntryField.get() == '' or passwordEntryField.get() == '':
            messagebox.showerror(
                'Error', 'All Fields Are Required', parent=rootN)

        else:
            f = open('setting.txt', 'w')
            f.write(fromEntryField.get()+','+passwordEntryField.get())
            f.close()
            messagebox.showinfo(
                'Information', 'CREDENTIALS SAVED SUCCESSFULLY', parent=rootN)

    rootN = Toplevel()
    rootN.title('Setting')
    rootN.geometry('650x340+350+90')
    rootN.resizable(False, False)

    rootN.config(bg='grey')

    Label(rootN, text='Your Information', image=logo, compound=LEFT, font=('goudy old style', 40, 'bold'),
          fg='black', bg='grey').grid(padx=60)

    fromLabelFrame = LabelFrame(rootN, text='From (Email Address)', font=('times new roman', 16, 'bold'), bd=5, fg='black',
                                bg='grey')
    fromLabelFrame.grid(row=1, column=0, pady=20)

    fromEntryField = Entry(fromLabelFrame, font=(
        'times new roman', 18, 'bold'), width=30)
    fromEntryField.grid(row=0, column=0)

    passwordLabelFrame = LabelFrame(rootN, text='Password', font=('times new roman', 16, 'bold'), bd=5,
                                    fg='black',
                                    bg='grey')
    passwordLabelFrame.grid(row=2, column=0, pady=20)

    passwordEntryField = Entry(passwordLabelFrame, font=(
        'times new roman', 18, 'bold'), width=30, show='*')
    passwordEntryField.grid(row=0, column=0)

    Button(rootN, text='SAVE', font=('times new roman', 18, 'bold'),
           cursor='hand2', bg='blue', fg='black', command=save).place(x=210, y=280)
    Button(rootN, text='CLEAR', font=('times new roman', 18, 'bold'),
           cursor='hand2', bg='blue', fg='black', command=clear1).place(x=340, y=280)

    f = open('setting.txt', 'r')
    for i in f:
        credentials = i.split(',')

    fromEntryField.insert(0, credentials[0])
    passwordEntryField.insert(0, credentials[1])

    rootN.mainloop()


def exit():
    result = messagebox.askyesno('Notification', 'Do you want to exit?')
    if result:
        root.destroy()
    else:
        pass


def clear():
    toEntryField.delete(0, END)
    subjectEntry.delete(0, END)
    text_area.delete(1.0, END)


root = Tk()
root.title('EMAIL APPLICATION GUI')
root.geometry('800x640+120+60')
root.config(bg='grey')
root.resizable(False, False)

Frame_Title = Frame(root, bg='grey')
Frame_Title.grid(row=0, column=0)
logo = PhotoImage(file='email.png')
Label_Title = Label(Frame_Title, text='  EMAIL APPLICATION ', image=logo, compound=LEFT, font=('luicda', 30, 'bold'),
                    bg='grey')
Label_Title.grid(row=0, column=0)
settingIMG = PhotoImage(file='setting.png')

Button(Frame_Title, image=settingIMG, bd=0, bg='grey', cursor='hand2',
       activebackground='grey', command=settings).grid(row=0, column=1, padx=20)

chooseFrame = Frame(root, bg='grey')
chooseFrame.grid(row=1, column=0, pady=10)
choice = StringVar()

SingleBut = Radiobutton(chooseFrame, text='One', font=('times new roman', 25, 'bold'), variable=choice, value='single', bg='grey', activebackground='white',
                        command=button_check)
SingleBut.grid(row=0, column=0, padx=20)

Mult_but = Radiobutton(chooseFrame, text='One+', font=('times new roman', 25, 'bold'), variable=choice, value='multiple', bg='grey', activebackground='white',
                       command=button_check)
Mult_but.grid(row=0, column=1, padx=20)

choice.set('single')

toLabelFrame = LabelFrame(root, text='To', font=(
    'times new roman', 16, 'bold'), bd=5, fg='white', bg='black')
toLabelFrame.grid(row=2, column=0, padx=30)

toEntryField = Entry(toLabelFrame, font=(
    'times new roman', 18, 'bold'), width=50)
toEntryField.grid(row=0, column=0)

browseImage = PhotoImage(file='browse.png')

browseButton = Button(toLabelFrame, text=' Browse', image=browseImage, compound=LEFT, font=('arial', 12, 'bold'),
                      cursor='hand2', bd=0, bg='grey', activebackground='grey', state=DISABLED, command=browse)
browseButton.grid(row=0, column=1, padx=20)

subjectLabelFrame = LabelFrame(root, text='Subject', font=(
    'times new roman', 16, 'bold'), bd=5, fg='white', bg='black')
subjectLabelFrame.grid(row=3, column=0, pady=10)

subjectEntry = Entry(subjectLabelFrame, font=(
    'times new roman', 18, 'bold'), width=62)
subjectEntry.grid(row=0, column=0)

emailLabelFrame = LabelFrame(root, text='Message', font=(
    'times new roman', 16, 'bold'), bd=5, fg='white', bg='black')
emailLabelFrame.grid(row=4, column=0, padx=20)

attachImage = PhotoImage(file='attachments.png')

Button(emailLabelFrame, image=attachImage, compound=LEFT, font=('arial', 12, 'bold'),
       cursor='hand2', bd=0, bg='grey', activebackground='grey', command=attachment).grid(row=4, column=2)

text_area = Text(emailLabelFrame, font=('times new roman', 14,), height=8)
text_area.grid(row=1, column=0, columnspan=2)


Button(root, text='Send', bd=0, bg='#3268a8', fg='black', font='lucida 17 bold',
       cursor='hand2', command=send_email).place(x=490, y=540)


Button(root, text='Clear', bd=0, bg='#3268a8', font='lucida 17 bold',
       cursor='hand2', command=clear).place(x=590, y=540)


Button(root, text='Exit', bd=0, bg='#3268a8', font='lucida 17 bold',
       cursor='hand2', command=exit).place(x=690, y=540)

total_Label = Label(root, font=('times new roman', 18, 'bold'),
                    bg='grey', fg='black').place(x=10, y=550)

sent_Label = Label(root, font=('times new roman', 18, 'bold'),
                   bg='grey', fg='black').place(x=100, y=550)

left_Label = Label(root, font=('times new roman', 18, 'bold'),
                   bg='grey', fg='black').place(x=180, y=550)

failed_Label = Label(root, font=('times new roman', 18, 'bold'),
                     bg='grey', fg='black').place(x=270, y=550)

root.mainloop()
