# Email-Automation

Email Marketing is one of the popular method to reach out customers. Emailing number of customers is time consuming and require manual workforce. To overcome this, I made this project to send mail to many persons at one go. 

The requirements to run the program is to have a Spreadsheet containing email address of the customers. 
The program is written in a way that it automatically extracts the spreadsheet id.        

At first, the program will extract the details in the spreadsheet and save it as a CSV/TSV file in local system. Then the CSV/TSV file is used to send emails using respective function. The program has been written to save the Extracted data as TSV file. If you want to save as CSV file, then change all the TSV specified in the program to CSV & set the delimiter as "," . specify the Output Directory path to save the file.

I have used argument parser to specify the spreadsheet link while running the programming itself. This enables it to run the program at single execution without asking input from the user.

It uses the SMTP server to send mails.specify the email address and Password of the sender from which you want to choose.

Workflow Outline:
             1. Extracting Spreadsheet ID from the spreadsheet link.
             2. Write the extracted data as TSV/CSV and store it in the local system.
             3. Extracting Email Address,Subject & Body of the mail from the TSV/CSV file.
             4. Using SMTP Server, sending mails to the customers.

Note :
            1. Make sure That the spreadsheet is "NOT RESTRICTED"
            2. Use the generated app Password as mail password. (Not the usual Mail ID password)
            3. Make sure 2-step verification is enabled.

How to run on vs code:
            Python "Script_name.py" -i "Spreadsheet Link"
             
             
