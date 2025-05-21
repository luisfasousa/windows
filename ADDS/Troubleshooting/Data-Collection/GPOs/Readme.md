Client machine data collection
Follow the steps below from a computer experiencing GPO application issues, from an elevated Command Prompt. If it's a user policy failing, make sure the logged-on user is one experiencing the problem.
Option A, Automatic collection - use Authscripts
Download the log collection scripts and unzip them, from here: http://aka.ms/authscripts
On the machine where the issue is reproduced, start PowerShell as Administrator
a. Navigate to location where you unzipped the script in step 1.
cd <the path of unzipped script from step 1>
b. Start the log collection by running:
.\start-auth.ps1
Wait for the script to finish initializing the log collection.
Reproduce the issue.
Stop the traces by running:
.\stop-auth.ps1
Zip the Authlogs folder that was created by the log collection script and upload it to the workspace.
