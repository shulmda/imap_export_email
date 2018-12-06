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
import datetime
import Queue
import Tkinter as tk
from tkinter import filedialog
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import string
import threading

# ToDo: The imap functions could be their own class and it should be possible to maintain the login without reconnect
# ToDo: Remove the print statements for Debugging
# ToDo: Command line interface?  Quiet Mode? Needed?
# ToDo: Remove Commented out print statments
# ToDo: Remove print statments or make a command line flag for when
# ToDo: Research: Do we really need to extend the OptionMenu? Why isn't there a method for this?
# ToDo: Make the number of threads user configurable?
# ToDo: Tests

valid_filename_chars = "-_.() %s%s" % (string.ascii_letters, string.digits)
threadcount = 5
invalidchars = {}

def check_char(c):
    """
        Function to check if a character is in the list of charecters allowed in filenames
    """
    # ToDo: Could be more efficient, it is looping through all valid chars each time.
    # ToDo: increase the valid characters.

    try:
        if c in valid_filename_chars:
            return c
        else:
            if c not in invalidchars:
                invalidchars[c] = 1
            else:
                invalidchars[c] = invalidchars[c] + 1
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

def get_folder_message_count(user_name,password,server,folder):
    """
        Function to count the number of messages in a folder
    """

    M = imaplib.IMAP4_SSL(server)
    M.login(user_name, password)
    rv, data = M.select(folder)
    if rv == 'OK':
        rv, data = M.search(None, "ALL")
        M.close()
        M.logout()
        if rv != 'OK':
            print "No messages found!"
            return 0

        messagelist = data[0].split()
        print "Message Count: %s " % (len(messagelist))
        return len(messagelist)
    else:
        print "ERROR: Unable to open mailbox ", rv
        M.close()
        M.logout()
        return 0

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
        message_time = datetime.datetime

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
    # ToDo: Fix issues where time is failing to be set.
    try:
        file_time = "%04d%02d%02d%02d%02d" % (
            message_time.year, message_time.month, message_time.day, message_time.hour, message_time.minute)

        # Timestamp the file to match the creation date of the email (making it easier to search)
        str_command = "touch -t %s \"%s\"" % (file_time, filename)
        os.system(str_command)

    except:
        print "Warning: Failed to set file time for email #:", num





def export_mailbox(M, output_directory, message_begin_num,message_end_num):
    """
        export all emails in the folder to files in output directory with names matching the subject.
    """
    count = 0
    # Loop through message list
    for num in range(message_begin_num,message_end_num + 1):

        try:
            #print "Export Message:", num
            export_message(M, num, output_directory)

        except Exception, e:
            print ('Failed to export message: ' + str(e))
            return -1

        count = count + 1



    return count



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

class ExportEmailThread(threading.Thread):
    def __init__(self, queue,stop_queue,email_account,password,imap_server,output_directory, email_folder,message_begin_num,message_end_num):

        threading.Thread.__init__(self)
        self.queue = queue
        self.stop_queue = stop_queue
        self.email_account = email_account
        self.password = password
        self.imap_server = imap_server
        self.output_directory = output_directory
        self.email_folder = email_folder
        self.message_begin_num = message_begin_num
        self.message_end_num = message_end_num
        self.batch_size= 100
    def run(self):
        self.export_folder()
        self.queue.put("done")



    def export_folder(self):
        print "email_account: %s" % self.email_account
        print "email_folder: %s " % self.email_folder

        current_message_number = self.message_begin_num

        end_message_number = 0

        retry_count = 0
        # Use the Batches to allow for fault tolerance
        while current_message_number < self.message_end_num:

            end_message_number = current_message_number + self.batch_size
            if (end_message_number > self.message_end_num):
                end_message_number = self.message_end_num

            if retry_count == 3:
                return -1

            # If a batch fails, restart the batch including login
            export_status = 0
            M = imaplib.IMAP4_SSL(self.imap_server)
            M.login(self.email_account, self.password)
            rv, data = M.select(self.email_folder)
            if rv == 'OK':
                #print "Processing mailbox: ", self.email_folder
                export_status = export_mailbox(M, self.output_directory, current_message_number,end_message_number)
                if export_status < 0:
                    # Increase Retry Count
                    retry_count = retry_count + 1

                    # ToDo: Make the retry delay a random interval
                    time.sleep(1)
                    continue
                else:

                    # Reset Retry count on successful export
                    retry_count = 0


                    # Transmit the number of emails export via the queue to main thread
                    self.queue.put("step:%s" % (export_status))

                    if current_message_number <= self.message_end_num:
                        current_message_number = end_message_number + 1

                # Check if a stop command has been issued (Note, it will only check on batch completion
                # ToDo: find an elegent way to stop without waiting for a batch to finish.
                try:
                    msg = self.stop_queue.get(0)
                    if (msg == "stop"):
                        M.close()
                        M.logout()
                        return
                except Queue.Empty:
                    continue


            else:
                print "ERROR: Unable to open mailbox ", rv

            M.close()
            M.logout()
        return



class GUI:
    def __init__(self, master):

        # ToDo: Rename the labels to have more useful names
        # ToDo: Add code to center the window
        # ToDo: Some folders with Unicode letters and spaces do not export

        self.master = master
        self.master.geometry("500x400")


        self.server = tk.StringVar(self.master, value='imap.gmail.com')
        self.labelserver = Label(self.master, text="Server:")
        self.labelserver.place(x=20, y=20)
        self.labelserverhelp = Label(self.master, text="(ie. imap.myserver.com)")
        self.labelserverhelp.place(x=340, y=20)

        self.entryserver = Entry(self.master, textvariable=self.server)
        self.entryserver.place(x=150, y=18)

        self.user_name = tk.StringVar(self.master, value='shulmda@gmail.com')
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
        self.child_thread_array = []
        self.message_count = 0
        self.processed_count = 0


    def tb_start(self):
        # ToDo: Wastes a bit of memory on restart by resetting the array instead of removing them all

        self.child_thread_array = []

        # ToDo: add code to add better error trapping
        if (self.destination.get() == ""):
            messagebox.showinfo("Export Error", "A Distination Directory must be selected")
            return

        if (self.selectedfolder.get() == "<Choose Mail Folder to export>"):
            messagebox.showinfo("Export Error", "A Mail Folder must be selected")
            return

        # Disable/Enable the buttons
        self.start_button.configure(state=DISABLED)
        self.stop_button.configure(state=NORMAL)

        self.prog_bar.step(-100)

        # create queue for messages from the child thread to the main thread (Progress Updates, Finished...)
        self.main_thread_queue = Queue.Queue()

        message_count = get_folder_message_count(self.user_name.get(), self.password.get(), self.server.get(),
                                                 self.selectedfolder.get())
        self.message_count = message_count




        if message_count > 0:

            start_num = 0
            end_num = 0

            threadcount_local = threadcount

            # Handle Case where the number of messages is less than the threadcount
            if (threadcount_local > message_count):
                threadcount_local = message_count


            thread_portion = int(round(float(message_count) / float(threadcount_local), 0))
            print "Messages per thread %s" % (thread_portion)

            # Loop through the threads and run one thread per segment
            for i in range(0, threadcount_local):
                print "Thread id: %s" % (i)

                # Add the remainder to the first thread end_num
                start_num = end_num + 1
                end_num = end_num + thread_portion

                if i == 0:
                    remainder = message_count - (thread_portion * threadcount_local)
                    end_num = end_num + remainder

                # create queue for messages from the main thread to each child thread (Stop/Pause message)
                child_thread_queue = Queue.Queue()
                self.child_thread_array.append(child_thread_queue)
                print "Starting Thread with range: %s to %s" % (start_num,end_num)
                if message_count > 0:
                    ExportEmailThread(self.main_thread_queue, child_thread_queue, self.user_name.get(), self.password.get(), self.server.get(), self.destination.get(), self.selectedfolder.get(), start_num, end_num).start()
                    self.master.after(100, self.process_mainthread_queue)

    def tb_stop(self):
        self.stop_button.configure(state=DISABLED)
        for queue in self.child_thread_array:
            queue.put("stop")


    def process_mainthread_queue(self):
        """
            Method to process the queue of messages from the child thread to the main thread
        """

        try:
            msg = self.main_thread_queue.get(0)

            # Process the messages from the child threads
            if ("step:" in msg):
                stepstr = msg[5:]
                stepamount = int(stepstr)
                interval = float(self.message_count) / 100
                already_processed = int(self.processed_count / interval)
                newly_processed = int((self.processed_count +  stepamount) / interval)

                # Step the progress bar only if the % has increased sufficiently
                if already_processed  < newly_processed:
                    self.prog_bar.step(newly_processed - already_processed  )

                self.processed_count = self.processed_count + stepamount


                self.labelcomplete['text'] = "%s%% Total: %s" % (round(float(self.processed_count) / float(self.message_count) * 100, 1), self.message_count)
                self.master.after(100, self.process_mainthread_queue)

            # ToDo: It reenables the Start Button before all threads are completed.  It should wait.
            if (msg == "done"):
                self.start_button.configure(state=NORMAL)
                self.stop_button.configure(state=DISABLED)

        except Queue.Empty:
            # If Queue us Empty then call itself after waiting 100ms
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


