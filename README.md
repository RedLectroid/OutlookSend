# OutlookSend
A C# tool to send emails through Outlook from the command line or in memory.  Designed to be used through execute-assembly in your favorite C2 Framework.  Tested with Cobalt Strike.

A simple tool to send emails through the locally installed instance of Outlook.  If Outlook is not running when the tool is used, it will launch Outlook minimized, send the email the kill the Outlook process.  If Outlook is already running it will simply use that instance.

There are 4 arguments it is expecting:

-s The Subject of the email.  Spaces are allowed in the Subject as long as it is wrapped in quotation marks "".
-a Attachment.  Will attach a local file to the email.  Please provide the full path, and quotation marks "" if there is a space in the path.
-r Recipient(s).  Comma seperated list of recipients if more than one.  Please use quotation marks "".
-b Body.  If 'HTML' flag is used, OutlookSend will look in the file 'emailBody.txt' for the HTML code.  Anything else will be send as a text email.

Ex.  OutlookSend2.exe -s "This is a test" -a C:\Users\ShellStorm\Desktop\payloads\NotMalicious.docm -r "shellstorm@notReal.email" -b HTML

There is very little error handling in this, so buyer beware.

Before compiling add the Microsoft Outlook 15.0 Object Library via Project->Add Reference->COM->Microsoft Outlook 15.0 Object Library
