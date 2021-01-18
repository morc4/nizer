# nizer
A couple of scrips that sync gmail attachments with drive, puting files into folders following a structure laid on a spreadsheet.

This program uses google oAuth, in order to use it you must create an account in google cloud console and create a json file of token credentials there in order to give access to the scripts for your gmail account.

Once you have the token, put it in the folder of the scripts and run the folder script, then go to your drive account, open the speadsheet created and create the names of the folders that you would like to have.

Once the spreadsheet is created, you can run the sync attachments scrip to syncronize everything. You must have python 3.6+ to run this script and the google oAuth libraries installed.
