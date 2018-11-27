#!/usr/bin/env python
#
# Python GUI to export emails into an IMAP folder with the filenames
# as the subject and the file date as the time the email was sent/received.
#
# DJS Nov 2018
#

import imaplib
import email
import email.header
import time
import os
from dateutil.parser import parse
import Queue
import Tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import string
import threading

# ToDo: The imap functions could be their own class and it should be possible to maintain the login without reconnected
# ToDo: If the connnection is closed, it should trap the error and auto-reconnect
# ToDo: Remove the print statements for Debugging


valid_filename_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)

class OptionMenu(tk.OptionMenu):
    """
        Extend the tkinter Options Menu to add the addOption method which doesn't seem to be present
    """

    def __init__(self, *args, **kw):
        self._command = kw.get("command")
        self.variable = args[1]
        tk.OptionMenu.__init__(self, *args, **kw)
    def addOption(self, label):
        self["menu"].add_command(label=label,
            command=tk._setit(self.variable, label, self._command))


def check_char(c):
    """
        Function to check if a character is in the list of charecters allowed in filenames
    """

    try:
        if c in valid_filename_chars:
            return c
        else:
            return ' '
    except:
        return ' '

def remove_invalid_filename_chars(value):
    """
        Function to remove any characters that are not allowed in a filename
    """
    # ToDo: This could be more efficient. I think that the value.encode is likely a waste of time. Also, it is over aggressive.
    # ToDo: Could remove some that are not invalid...

    try:
        filename = ""
        try:
            s = value.encode('ascii', 'ignore')
        except:
            s = value
        for c in s:
            filename += check_char(c)
        return filename

    except:
        print "Error converting Subject!"
        return ""




def list_folders(user_name,password,server):
    """
        Function to list the imap folders
    """

    # Todo: Could be a bit more efficient
    folder_list = []
    try:

        M = imaplib.IMAP4_SSL(server)
        M.login(user_name, password)
        selectone = False
        for i in M.list()[1]:
            l = i.decode().split(' "/" ')

            folder = "%s" %(l[1])
            folderstr = folder[1:-1]
            folder_list.append(folderstr)

        M.logout()
        return folder_list
    except:
        return []


def export_message(M,num,output_directory):

    """
        Function to list the imap folders
    """

    # ToDo: The setting of the file time might fail. Should verify that it worked using iostat after setting and perhaps retry
    rv, message = M.fetch(num, '(RFC822)')

    msg = email.message_from_string(message[0][1])
    subj_decode = email.header.decode_header(msg['Subject'])[0]
    subject = remove_invalid_filename_chars(subj_decode[0])
    if (subject == ''):
        print "Empty Subject on email %s " % num
    # print "subject: %s" %(subject)
    if len(subject) > 200:
        subject = subject[:200]
    timestamp_decode = email.header.decode_header(msg['Date'])[0]
    timestamp = unicode(timestamp_decode[0])
    # print "timestamp: %s" %(timestamp)
    try:
        message_time = parse(timestamp)
    except:
        message_time = parse("Thu Nov 22 00:00:00 +0200")

    if rv != 'OK':
        print "ERROR getting message", num
        return
    # print "Writing message ", num
    if (subject == ''):
        filename = '%s/No Subject (%s).eml' % (output_directory,  num)
    else:
        filename = '%s/%s (%s).eml' % (output_directory, subject, num)

    f = open(filename, 'wb')
    f.write(message[0][1])
    f.close()
    # YYYYMMDDhhmm
    file_time = "%04d%02d%02d%02d%02d" % (
    message_time.year, message_time.month, message_time.day, message_time.hour, message_time.minute)

    str_command = "touch -t %s \"%s\"" % (file_time, filename)
    # print "str_command: %s " % (str_command)
    os.system(str_command)
    #stinfo = os.stat(filename)
    # print stinfo



def export_mailbox(M, output_directory, skip, queue, stop_queue):
    """
        export all emails in the folder to files in output directory with names matching the subject.
    """

    # ToDo: MessageBox should be handled in the GUI Thread, not here.

    rv, data = M.search(None, "ALL")
    if rv != 'OK':
        print "No messages found!"
        return

    messagelist = data[0].split()
    print "Message Count: %s " %(len(messagelist))
    count = 0
    interval = len(messagelist) / 100
    if interval == 0:
        interval = 1
    #print "interval: %s" % interval
    skip_count = int(skip)
    print "skip_count: %s" % skip_count
    #skip_count= 0
    for num in messagelist:
        #print "count: %s" % count

        if count < skip_count:
            count = count + 1
            if (count % interval == 0):
                print "Skipped %s%% (%s of %s)" % (
                (round(float(count) / float(len(messagelist)) * 100, 1)), count, len(messagelist))
                queue.put("step:%s,%s" % (count, len(messagelist)))
            continue
        else:
            try:
                #print "Export: %s" % num
                export_message(M, num, output_directory)

            except:
                queue.put("step:%s,%s" % (count -1, len(messagelist)))
                time.sleep(1)
                queue.put("done")
                messagebox.showinfo("Export Error", "Export Error. Press 'Start Export' to restart the process from where it left off.")
                return
            count = count + 1

        if (count % interval == 0 ):
            print "Exported %s%% (%s of %s)" % ((round(float(count)/float(len(messagelist) ) * 100,1)), count,len(messagelist) )
            queue.put("step:%s,%s" % (count, len(messagelist)))

        # Check if a stop command has been issued
        try:
            msg = stop_queue.get(0)
            if (msg == "stop"):
                queue.put("step:%s,%s" % (count, len(messagelist)))
                time.sleep(1)
                queue.put("done")
                return
        except Queue.Empty:
            continue

    queue.put("step:%s,%s" % (count, len(messagelist)))
    print "Exported Emails: %s" % (count)

class ExportEmailThread(threading.Thread):
    def __init__(self, queue,stop_queue,email_account,password,imap_server,output_directory, email_folder,skipcount):
        threading.Thread.__init__(self)
        self.queue = queue
        self.stop_queue = stop_queue
        self.email_account = email_account
        self.password = password
        self.imap_server = imap_server
        self.output_directory = output_directory
        self.email_folder = email_folder
        self.skipcount = skipcount

    def run(self):
        self.export_folder()
        self.queue.put("done")



    def export_folder(self):
        print "email_account: %s" % self.email_account
        print "email_folder: %s " % self.email_folder
        M = imaplib.IMAP4_SSL(self.imap_server)
        M.login(self.email_account, self.password)
        rv, data = M.select(self.email_folder)
        if rv == 'OK':
            print "Processing mailbox: ", self.email_folder
            export_mailbox(M, self.output_directory, self.skipcount, self.queue, self.stop_queue)
            M.close()
        else:
            print "ERROR: Unable to open mailbox ", rv
        M.logout()




class GUI:
    def __init__(self, master):

        # ToDo: Rename the labels to have more useful names
        # ToDo: Add code to center the window
        # ToDo: Some folders with Unicode letters and spaces do not export

        self.master = master
        self.master.geometry("500x400")


        self.server = tk.StringVar(self.master, value='')
        self.labelserver = Label(self.master, text="Server:")
        self.labelserver.place(x=20, y=20)
        self.labelserverhelp = Label(self.master, text="(ie. imap.myserver.com)")
        self.labelserverhelp.place(x=340, y=20)

        self.entryserver = Entry(self.master, textvariable=self.server)
        self.entryserver.place(x=150, y=18)

        self.user_name = tk.StringVar(self.master, value='')
        self.labelusername = Label(self.master, text="User Name:")
        self.labelusername.place(x=20, y=50)
        self.entryusername = Entry(self.master, textvariable=self.user_name)
        self.entryusername.place(x=150, y=48)

        self.labelpassword = Label(self.master, text="Password:")
        self.labelpassword.place(x=20, y=80)
        self.password = tk.StringVar(self.master, value='')
        self.entrypassword = Entry(self.master, textvariable=self.password, show="*")
        self.entrypassword.place(x=150, y=78)

        self.btn_login = Button(self.master, text="Login",command=self.update_folders_options)
        self.btn_login.place(x=150, y=110)


        self.destination = tk.StringVar(self.master, value='')
        self.labeldestination = Label(self.master, text="Destination Folder:")
        self.labeldestination.place(x=20, y=140)
        self.entrydistinationdir = Entry(self.master, textvariable=self.destination, state='readonly')
        self.entrydistinationdir.place(x=150, y=138)
        self.btn_browse = Button(self.master,text="Browse", command=self.find_dir)
        self.btn_browse.place(x=350, y=138)

        self.labelprogress = Label(self.master, text="Progress:")
        self.labelprogress.place(x=20, y=230)
        self.prog_bar = ttk.Progressbar(
            self.master, orient="horizontal",
            length=200, mode="determinate"
            )

        self.prog_bar.place(x=150, y=230)

        self.skip = tk.StringVar(self.master, value='0')
        self.labelmessageindex = Label(self.master, text="Message Index:")
        self.labelmessageindex.place(x=20, y=260)
        self.labelserverhelp = Label(self.master, text="(ie. message # offset)")
        self.labelserverhelp.place(x=340, y=260)

        self.entryskip = Entry(self.master, textvariable=self.skip)
        self.entryskip.place(x=150, y=258)


        self.labelpctcomplete = Label(self.master, text="Percent Complete:")
        self.labelpctcomplete.place(x=20, y=200)
        self.labelcomplete = Label(self.master, text="0%")
        self.labelcomplete.place(x=150, y=200)


        self.labelmailfolder = Label(self.master, text="Mail Folder:")
        self.labelmailfolder.place(x=20, y=170)
        self.selectedfolder = tk.StringVar(self.master)
        self.selectedfolder.set("<Choose Mail Folder to export>")
        self.optionMenu = OptionMenu(self.master, self.selectedfolder, '<Choose Mail Folder to export>')
        self.optionMenu.place(x=148, y=168)



        self.start_button = Button(self.master, text="Start Export",state=DISABLED, command=self.tb_start)
        self.start_button.place(x=150, y=300)

        self.stop_button = Button(self.master, text="Stop Export",state=DISABLED, command=self.tb_stop)
        self.stop_button.place(x=150, y=325)


    def tb_start(self):

        # ToDo: add code to add better error trapping
        if (self.destination.get() == ""):
            messagebox.showinfo("Export Error", "A Distination Directory must be selected")
            return

        if (self.selectedfolder.get() == "<Choose Mail Folder to export>"):
            messagebox.showinfo("Export Error", "A Mail Folder must be selected")
            return

        self.start_button.configure(state=DISABLED)
        self.stop_button.configure(state=NORMAL)

        # create 2 queues.
        # 1 for messages from the child thread to the main thread (Progress Updates, Finished...)
        # 1 for messages from the main thread to the child thread (Stop/Pause message)
        self.main_thread_queue = Queue.Queue()
        self.child_thread_queue = Queue.Queue()

        ExportEmailThread(self.main_thread_queue, self.child_thread_queue, self.user_name.get(), self.password.get(), self.server.get(), self.destination.get(), self.selectedfolder.get(), self.skip.get()).start()
        self.master.after(100, self.process_mainthread_queue)

    def tb_stop(self):

        self.child_thread_queue.put("stop")

    def process_mainthread_queue(self):
        """
            Method to process the queue of messages from the child thread to the main thread
        """

        try:
            msg = self.main_thread_queue.get(0)

            # Process the messages from the child thread
            if ("step:" in msg):
                stepstr = msg[5:]
                stepamount,steptotal = stepstr.split(",")
                self.skip.set(stepamount)
                self.prog_bar.step(1)
                self.labelcomplete['text'] = "%s%% Total: %s" % (round(float(stepamount) / float(steptotal) * 100, 1), steptotal)
                self.master.after(100, self.process_mainthread_queue)

            if (msg == "done"):
                self.start_button.configure(state=NORMAL)
                self.stop_button.configure(state=DISABLED)

        except Queue.Empty:
            self.master.after(100, self.process_mainthread_queue)

    def update_folders_options(self):
        """
            Method to get the list of folders and populate the options Menu
        """
        folder_list = list_folders(self.user_name.get(), self.password.get(), self.server.get())
        if (len(folder_list) == 0):
            messagebox.showinfo("Login Failed", "Login Failed. Please Check Credentials")
        else:
            menu = self.optionMenu.children["menu"]
            menu.delete(0, 'end')
            selectone = False
            for folder in folder_list:
                self.optionMenu.addOption(folder)
                if (selectone == False):
                    self.selectedfolder.set(folder)
                    selectone = True
            self.start_button.configure(state=NORMAL)


    def find_dir(self):
        """
            Method to open a Directory selection dialog
        """

        # ToDo: Figure out why the warning "Class FIFinderSyncExtensionHost is implemented in both" is appearing, seems to be a MacOs issue
        file_path = filedialog.askdirectory(title="Select Destination")
        self.destination.set(file_path)

root = Tk()
root.title("Email Export")
main_ui = GUI(root)
root.mainloop()


