# Created by Dr Tim Wilding,  2022
# Copyright (c) 2022, Dr Tim Wilding
# All rights reserved.
#
# This source code is licensed under the BSD-style license found in the
# LICENSE file in the root directory of this source tree.

import win32com.client
import pandas as pd
from pathlib import Path
from tkinter import filedialog
import tkinter
import atexit
import re
import sys
import pythoncom


def substitute_variables(in_text, row_var):
    # return re.sub(r'(\$[a-zA-Z0-9_-]+)', my_replace, in_text)
    return re.sub(r'(\$[a-zA-Z0-9_-]+)', lambda match_obj: str(row_var[match_obj.group(0).split('$')[1]]), in_text)


def parse_excel_file(in_filename, action):
    was_error = 0
    try:
        file = pd.ExcelFile(
            in_filename)  # Establishes the excel file you wish to import into Pandas
    except PermissionError:
        print("Could not open the file, try close it in excel and check permissions before trying again.\n")
        sys.exit(1)
    except AssertionError:
        print("Please select a file.\n")
        sys.exit(1)

    fullpath = Path(in_filename)
    filepath = str(fullpath.parent)

    df = file.parse('Sheet1')  # Uploads Sheet1 from the Excel file into a dataframe

    # globalise the row variable because I cannot figure out how to pass it to the my_replace function
    #global row

    for index, row in df.iterrows():  # Loops through each row in the dataframe
        email = (row['Email'])  # Sets dataframe variable, 'email' to cells in column 'Email Addreses'
        # subject = (row['Subject'])  # Sets dataframe variable, 'subject' to cells in column 'Subject'
        # body = str((row['Email HTML Body']))  # Sets dataframe variable, 'body' to cells in column 'Email HTML Body'
        subject = substitute_variables(str(row['Subject']), row)
        # now we replace all ${whatever} in the message with the data from the row with that column name
        body = substitute_variables(str(row['Message']), row)

        if pd.isnull(email):  # Skips over rows where there is no email set
            continue

        olMailItem = 0x0  # Initiates the mail item object
        obj = win32com.client.Dispatch("Outlook.Application")  # Initiates the Outlook application
        newMail = obj.CreateItem(olMailItem)  # Creates an Outlook mail item
        newMail.To = email  # Sets the mail's To email address to the 'email' variable
        newMail.Subject = subject  # Sets the mail's subject to the 'subject' variable
        newMail.HTMLbody = (r"" +
                            body +
                            "")  # Sets the mail's body to 'body' variable

        feed_back_file = (row['Feedback File'])

        if not pd.isnull(feed_back_file):
            for attach_name in str(feed_back_file).split(','):
                file_attachment_name = filepath + '\\' + substitute_variables(attach_name, row)
                try:
                    newMail.Attachments.Add(Source=file_attachment_name)
                except pythoncom.com_error:
                    print("There was a problem with attaching the file %s.  \nCheck the file exists and is not open or "
                          "read only." % file_attachment_name)
                    # as there has been an error change action to do nothing, the rest will be checked but nothing
                    # further will be sent, or done
                    print("Switched to check only mode, no emails will be sent, created or saved, "
                          "attachments will continue to be checked.")
                    action = 0
                    was_error = 1
        if action == 0:
            pass
        elif action == 1:
            newMail.Display(False)
        elif action == 2:
            newMail.save()
        elif action == 3:
            newMail.send

        # newMail.display()  # Displays the mail as a draft email use (True) to wait holding
        # script before window closes
        # this will send it
        # newMail.send
        # this will save it
        # newMail.save()
    return was_error


if __name__ == '__main__':
    atexit.register(input, "Enter any Key to Close/Exit")

    root = tkinter.Tk()
    root.withdraw()

    filename = filedialog.askopenfilename(title="Spreadsheet results file",
                                          filetypes=(("xlsx", '*.xlsx'), ("all", "*.*")))

    # first we pass through everything and display the emails only to check for errors. if no errors then create mails
    if parse_excel_file(in_filename=filename, action=0) == 0:
        parse_excel_file(in_filename=filename, action=2)
        print("Check your draft email folder for the created emails.")
        if input("Do you want to send the emails: (enter yes to send)").lower().strip() == "yes":
            # here we will actually send them
            print("Ok sending the emails, you will need to manually delete the draft emails from the draft folder.")
            parse_excel_file(in_filename=filename, action=3)
    else:
        print("Checking process failed no emails were saved,sent or created.  Check output for errors.")




