AutomatePingCastle
===
This simple script is designed to enhance the efficiency of running PingCastle, a comprehensive tool for Active Directory Security Assessment.

The script will perform the following:
1. **Reports Folder Creation:** Automatically generates a 'Reports' folder within the main PingCastle directory.
2. **Report Storage:** Saves the PingCastle HTML report in the 'Reports' folder for easy access and organization.
3. **Change Detection:** Compares the current scan's XML data file with the previous one to identify any changes since the last PingCastle scan.
4. **Email Notifications:** Sends an email through a specified SMTP server to a recipient of your choosing. 
5. **Change Alerts:** If there are any differences detected from the last scan, the email subject line will highlight that there are changes to the recepient.
6. **Automatic Updates:** Executes PingCastle's update process to ensure the tool remains current with the latest features.

## Usage

Follow these steps to set up and run the AutomatePingCastle script:

1. **Download PingCastle:** Visit [PingCastle's download page](https://www.pingcastle.com/download/) and download the tool.
2. **Prepare the Environment:** Unzip the downloaded file and rename the folder to 'PingCastle'.
3. **Script Placement:** Ensure that the AutomatePingCastle script is located in the same directory as the 'PingCastle' folder. _It should not be inside the 'PingCastle' folder!_
4. **SMTP Configuration:** Enter details for an SMTP Server, including Username, Password, and Sender Email.
5. **Execute the Script:** Run the script to start automating your PingCastle scans.