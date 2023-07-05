# SharePoint-AutoArchiving
<p>This PowerShell script automatically manages the creation of archive folders on a SharePoint site each month and sends email notifications regarding the process.</p>
<h3>Description</h3>
<p>This script is designed to automate the process of creating monthly archive folders in a SharePoint site, updating a settings list, and sending email notifications regarding the success or failure of these operations. This can be particularly useful for SharePoint administrators or users who need to manage archives on a regular basis.</p>
<p>
  The script:
  <ol>
    <li>Loads the SharePoint PowerShell snap-in.</li>
    <li>Accesses the SharePoint site and retrieves a collection of all lists within a subsite.</li>
    <li>Retrieves the current year and month to construct the names of the archive folders to be created.</li>
    <li>Checks if the archive folders exist and creates them if they do not.</li>
    <li>Updates the "Destination Document Library" in the "Archive Settings" list.</li>
    <li>Sends an email notification detailing whether the archive folders were created successfully or not.</li>
    <li>Disposes of the connections to the SharePoint site to prevent memory leaks.</li>
  </ol>
</p>
<h3>Configuring the Script</h3>
<p>
  The script contains several variables that may need to be configured to match your SharePoint environment:
  <ul>
    <li><b>'$ordersArchiveUrl'</b>: The URL of the SharePoint site where the archive folders should be created.</li>
    <li><b>'Get-SPSite'</b>: The base URL of your SharePoint site collection.</li>
    <li><b>'Get-SPWeb'</b>: The URL of the subsite where the archive folders should be created.</li>
    <li><b>'$ordersUrl'</b>: The URL of the site containing the "Archive Settings" list that should be updated.</li>
    <li><b>'Send-MailMessage'</b>: The details of the SMTP server and email addresses for the email notifications.</li>
  </ul>
</p>
<h3>Usage</h3>
<ol>
    <li>Open PowerShell.</li>
    <li>Navigate to the directory containing the <b>'CreateMonthlyArchiveFolders.ps1'</b> script.</li>
    <li>Run the script with the command <b>'.\CreateMonthlyArchiveFolders.ps1'</b>.</li>
  </ol>
This will execute the script and start the process of creating the archive folders and sending the email notifications.
<br><br>
Please note that this script is designed to run with the current date. If you want to use it for a different month, you'll need to modify the date parameters in the script.
</p>
<h3>Contributing</h3>
<p>Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.</p>
<h4>Authors</h4>
    <p>ch11nv11n
    <br>
    Your Contact Information
    <ul><li><b>DISCORD</b>: ch11nv11n</li></ul>
    </p>
<h4>License</h4>
<p>This project is licensed under the MIT License - see the LICENSE file for details.</p>
