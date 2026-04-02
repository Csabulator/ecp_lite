# ecp_lite
**Overview**

This script provides a web-based control panel for managing Exchange Remote Mailboxes, Shared Mailboxes, Distribution Groups, and Room/Equipment Mailboxes. It runs a simple HTTP server on your local machine, allowing you to perform common Exchange recipient management tasks via a web browser.

**Disclaimer**
- The script is provided as-is and it is not recommended to use it in production environments without proper testing and security considerations. I cannot be held responsible for any damage or issues caused by using this script. Use it at your own risk.
- If you would like to see improvements or have any suggestions, please feel free to contribute to the repository.
#
**1. Prerequisites**
- Windows Server or Workstation with Exchange Management Tools installed.
- Exchange Management Shell (EMS) must be available.
- PowerShell (preferably 5.1 or later).
- Administrator rights on the machine.
- .NET Framework (required for System.Net.HttpListener).
#
**2. Preparation**

**a. Exchange Management Shell**
- Ensure you can run Exchange cmdlets like Get-RemoteMailbox, Set-RemoteMailbox, Get-DistributionGroup, etc.
- If not, install the Exchange Management Tools for your environment.

**b. Script Placement**
- Download the latest version of the script.
- Place the script file in a secure folder on your workstation.
#
**3. Running the Script**
- Open PowerShell as Administrator.
- Navigate to the script's folder:
- cd "C:\Path\To\Your\Script"
- Run the script:
- ecp_lite.ps1
- Wait for the message:
- Exchange Control Panel running on http://localhost:8080/
#
**4. Accessing the Web Interface**
- Open your web browser (preferably Edge).
- Go to: http://localhost:8080/
- You will see the Exchange Control Panel home page.
#
**5. Main Features**

**a. Navigation**
- List All Mailboxes: View all remote mailboxes (paged, 50 per page).
- List All Shared Mailboxes: View all shared remote mailboxes.
- List All Distribution Groups: View all distribution groups.
- List All Room/Equipment Mailboxes: View all room/equipment mailboxes.
- Search for mailboxes, distribution groups, or room/equipment mailboxes.
- Enable Remote Mailbox: Create a new remote mailbox.

**b. Paging**
- Lists are paged (50 items per page).
- Use Previous and Next buttons to navigate.
#
**6. Managing Mailboxes**

**a. Viewing Mailbox Details**
- Click a mailbox name to view details.
- See properties like Name, Display Name, Alias, UPN, Primary Email, Email Addresses, Remote Routing Address, Mailbox Type, Archive info, etc.

**b. Updating Mailbox**
- Edit Primary Email Address or Email Addresses.
- Click Update Mailbox to save changes.
- If you change the Primary Email Address, the Remote Routing Address will update automatically.
- Email Address rules:
- Format: smtp:xyz@yourdomain.com;
- Always separate email addresses with ; and a space after it.

**c. Disabling a Mailbox**
- Click Disable Mailbox.
- You will be prompted for confirmation.
- Confirm to remove Exchange properties from the user (cannot be undone easily).

**d. Converting Mailbox Type**
- Use Convert to Shared or Convert to User to change mailbox type.
- Note: Shared mailbox delegation works in Exchange Online only.
  
**e. Disabling Archive**
- If the mailbox has an archive, a Disable Archive button appears.
- Click to disable the archive.
  #
**7. Managing Distribution Groups**

**a. Viewing Distribution Group Details**
- Click a group name to view details and members (sorted alphabetically).

**b. Disabling a Distribution Group**
- Click Disable Distribution Group.
- Confirm the action.
- The group will be disabled and Exchange properties are removed. Object will remain in Active Directory.

**c. Deleting a Distribution Group**
- Click Delete Distribution Group.
- Confirm the action.
- The group will be deleted from Active Directory.
#
**8. Managing Room/Equipment Mailboxes**

**a. Viewing Details**
- Click a room/equipment mailbox to view details.

**b. Disabling or Deleting**
- Use Disable Room/Equipment Mailbox or Delete Room/Equipment Mailbox.
- Each action requires confirmation.
#
**9. Creating a New Remote Mailbox**
- On the home page, use the Enable Remote Mailbox form.
- Enter the alias (username) and click Create.
- Use Active Directory to check UserPrincipalName of the user to avoid typos.
- The script will create a new remote mailbox with the correct routing addresses.
- The mailbox will be a user mailbox. If a Shared mailbox is needed, convert it afterwards.
#
**10. Setting the OwnerN Attribute**
- On the home page, use the Set OwnerN Attribute form.
- Enter the account (UPN or sAMAccountName) and the desired OwnerN (UPN).
- Only accounts in specific OUs (as defined in the script) can be modified.
- If successful, the attribute is set; otherwise, an error or restriction message is shown.
#
**11. Searching**
- Use the search forms on the home page to find mailboxes, distribution groups, or room/equipment mailboxes by name.
- Results are paged.
#
**12. Stopping the Script**
- To stop the web server, press Ctrl+C in the PowerShell window.
- If not working, close the PowerShell window.
#
**13. Security Notes**
- The server listens only on localhost (not accessible remotely).
- Only users with sufficient Exchange permissions can make changes.
- Do not expose this script to the internet.
#
**14. Troubleshooting**
- Port 8080 in use: Change the port in the script if needed.
- Exchange cmdlets not found: Ensure Exchange Management Shell is installed and available.
- Permission errors: Run PowerShell as Administrator and ensure your account has Exchange recipient management rights.
- Script errors: Review the PowerShell console for error messages.
- For any operation, simply follow the on-screen instructions and use the navigation buttons provided.
