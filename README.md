# GoogleCalendar_To_ExcelSheets
A python script that converts data from the google calendar to an excel custom format.

Read the api modules to install.txt for dependecies setup.

You will need to register the app in the google developers console and allow the APIs to be called on the project.

The beggining run (on a new user) is a bit tideous but only needs to be run once. You need this to generate your credentials which will be saved in a pickle file (then you can comment back the "first run on a new user part")

You will need to generate credentials for the calendar you want to use in the google web developer console. In order to do that you will need to firt run the "first run on a new user part". This will give you a link to open in the web browser where you have to log in with your google account that has the calendar. This will give you a token as a .json file that you need to add to your project directory. Then you can comment back the "first run on a new user part" as the token is saved as a token.pkl file which is loaded at each run.

Disclaimer! The process might change as I am not keeping this project up to date with whatever changes google might implement.
