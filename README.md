# ScheduleMeets
This code was created by Clare Ciriaco For Discovery College. We are putting it up here in the hope it can help other schools shedule online video conferences based on their timetabled classes automatically 
You will need to create a google account to allow this script to run and create calendar events on its own calendar. using your own account is not recommended.
From this new account you can open google drive and create a spreadsheet.
In this spreadsheet under tools open the script editor copy in the code from onlinelearning.gs.
Set up your database connection to your school management system
Create the queries to pull the data fro your school database - for our system we have classes that are taught combined we have identified and merged these classes in this script
set up a trigger to run the script daily. Trying to do more than a day's worth of lessons causes timeouts.
