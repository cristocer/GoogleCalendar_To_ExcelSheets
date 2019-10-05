# GoogleCalendar_To_ExcelSheets
A python script that converts data from the google calendar to an excel custom format.

Scope: I needed a way to export the google calendar data of a company I worked for, to Excel format, and adjusted to my manager preference. There are version of this project on the internet provided by some companies, but 1. They don't work properly. 2. They also cost money(like....). 3. You can barely adjust the column and row format or how the data will be collected and identified.

Calendar_to_Excel is the "perfect" complete version of the project.
Calendar_to_ExcelV2 is an almost complete version of the project as it has a bonus additional sheet with the operational plan that I didn't manage to finish to fully test.

I modified the original code as I don't want to disclose sensitive and personal information about the company and members I worked for. Enjoy the code and if you want to sell a version of this project I would appreciate a reference or job opportunity  :)


Disclaimer! The process might change as I am not keeping this project up to date with whatever changes google might implement.

Read the api modules to install.txt for dependecies setup.

You will need to register the app in the google developers console and allow the APIs to be called on the project.

The beginning  run (on a new user) is a bit tedious but only needs to be run once. You need this to generate your credentials which will be saved in a pickle file (then you can comment back the "first run on a new user part")

You will need to generate credentials for the calendar you want to use in the google web developer console. In order to do that you will need to first run the "first run on a new user part". This will give you a link to open in the web browser where you have to log in with your google account that has the calendar. This will give you a token as a .json file that you need to add to your project directory. Then you can comment back the "first run on a new user part" as the token is saved as a token.pkl file which is loaded at each run.




