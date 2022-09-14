This script, when run, opens a file selector dialog.  The user should select the student result spreadsheet file.  An example file format is included.  The script runs through generating emails much like a mailmerge.  Emails can be generated from other column data within each row (eg name/results etc).  The advantage of using this script over standard mailmerge is that one can attach files and the file can be a different file for each student.  This is handy for sending out results to students when you have a directory of feedback files.  You need the filenames to be unique to each student and the names are put in the spreadsheet (using another column if required, eg student number or name).

The script first runs through creating an email which is saved in your outlook draft folder.  Then pauses whilst you check through them for errors.   You can then abort, or choose to have the script send the emails.  You will then need to manually delete the drafts from the draft email box.

Use with caution as you can accidentally send junk, or unintended messages to anyone.

You need outlook install for this to work.  If you have multiple profiles then ensure you have outlook opened with the profile you actually want to use to send the emails.

Before running install the requirements from the requirements file

pip install -r requirements.txt
