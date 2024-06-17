## M365 Teams Policy Update 
User guide: [PDF](https://github.com/ITAutomator/M365-Teams-Policy-Update/blob/main/M365%20Teams%20Policy%20Update%20Readme.pdf)   
Download: [ZIP](https://github.com/ITAutomator/M365-Teams-Policy-Update/archive/refs/heads/main.zip)   
Website: [WWW](https://www.itautomator.com/m365-teams-policy-report-and-update/)  
(or click the green *Code* button (above) and click *Download Zip*)    

**Overview**   
Reports and updates Teams External Access Policies that aren't (yet) visible in the admin center. 
![image](https://github.com/ITAutomator/M365-Teams-Policy-Update/assets/135157036/ac157400-abac-4ef6-8813-42bbed6b5fb8)




**Usage**   
Use this code in 2 phases to create a CSV report of the editable properties of your users in Entra. 

How it works
Use this code in 2 phases to report on, then update the Teams ExternalAccessPolicy for your users.
The assumption is that, in general Teams chats to outside companies and individuals is blocked.  The code checks a certain group of users and switches them to allowed.  All other users are switched to blocked.

Phase 1: Report
Run the M365TeamsPolicyReport.ps1 (or .cmd) and enter your admin credentials.
This will output a CSV file containing your users and the policies they have been assigned.
•	Only Enabled accounts are reported.  Only members are reported (vs guests).
•	By default, no policy might be assigned (aka Global or <none>).  
•	For Teams users, the default policy seems to be FederationAndPICDefault which allows external access at the user level. (see below)


[See User Guide for more]
