# FedEx-Automation

This repo contains many automation programs for FedEx and other related tools.

IMPORTANT: You cannot connect to databases if working from YYZHQ / connected to a YYZHQ network. The only alternative is to disconnect, connect to a mobile hotspot, and connect to zscaler. You should then be able to run theses programs.

IMPORTANT: Make sure to install oracle instantclient and unzip it under: 'main_automation_programs/support-files/Oracle/'          folder -> e.g 'main_automation_programs/support-files/Oracle/your-instantclient-install'

Download from here: https://www.oracle.com/ca-en/database/technologies/instant-client/downloads.html

Go to the latest version and make sure to get the BASIC version, not light. Unzip at folder specified above.

VERY IMPORTANT ENV VARIABLES INFORMATION:
Many of these programs rely on .env files to store login / status variables.
Please make sure to include an .env file in BOTH the main-automation-programs directory and the HVSDistribution directory.

For the main-automation-programs directory, please include these variables:
misa-username =  first letter of your first name, first letter of your last name, emp id -> E.G for someone with the name john smith and emp id 1235 -> 'js12345'
fedex-email = 'your-fedex-email@fedex.com'
gmail-email = 'your-gmail-email@gmail.com'
gmail-app-password = 'your gmail app password' -> How to create a gmail app password: https://www.youtube.com/watch?v=GsXyF5Zb5UY
pw = 'your-fedex-login-password'
last_fta_filepath='lastest/fta/filepath' #Please navigate to 'main-automation-programs/reports/audits/fta-corrections/' to see the latest available fta file. It should be something like: 'main-automation-programs/reports/audits/fta-corrections/May2025/FTA Corrections_May14_2025.xlsm'. Find the latest fta file and copy its relative address to this variable.

For the HVSDistribution directory, please create an .env file and include these variables (also keep the .env file in the same directory as an executable, if you export the program into an executable)

okta_username='your-fedex-emp-id'
okta_password='your-fedex-password'
okta_automation='1'

You won't be able to pull / push or do anything with git while connected to a YYZHQ network. Make sure you are working from a vpn before using git.

Do do a pip install for all the libraries listed in the libraries.txt file