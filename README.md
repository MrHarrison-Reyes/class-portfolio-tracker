# class-portfolio-tracker
This tracker (with Google Sheets Template) creates an attendance and competency tracker for students and teachers. The tracker will send an automated monthly email to parents and students to inform them of progress and attendance rates for the month.

Before creating the AppScript, your Google Sheet must consist of the following tabs:
 * Class Attendance Tracker
 * Math Portfolio Progress
 * Then a sheet for each individual student with student email and parent email
 * I have created a template here: https://docs.google.com/spreadsheets/d/1esrAGQllzy7zVb60ldvzii7Ay4WwL5G21e3FoyIadhQ/edit?usp=sharing

To send automated emails follow these steps in the appscript editor:
   1. On the left select "Triggers, symbol looks like a clock
   2. Bottom right of screen, click "+ Add Trigger"
   3. "Choose Which Function to Run:" Select "sendMonthlyEmails"
   4. "Choose which deployment should run:" Select "Head"
   5. "Select Event Source" select "Time-Driven"
   6. "Select type of time based trigger", select "Month Timer"
   7. "Select Day of Month", select which day you prefer, I prefer 28th of each month to garantee all months during school year are sent.
   8. "Select Time of Day", select which time you prefer, I would suggest 4pm or 7AM
   9. "Failure Notification Settings", select "immediately"
   10. Hit "Save" and give it permissions.
   You are all set!
