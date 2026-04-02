<#
.DESCRIPTION
This script is used to create a simple web-based control panel for managing Exchange Online recipients such as Remote Mailboxes, Distribution Groups, and Room/Equipment Mailboxes. It uses the built-in HttpListener class in PowerShell to serve HTML pages and handle user interactions.
It is created for those administrators who has no prior PowerShell experience but need to manage Exchange Online recipients. It can also be used by experienced administrators as a quick and easy way to perform common recipient management tasks without having to remember and type PowerShell cmdlets.

DISCLAIMER:
- The script is provided as-is and it is not recommended to use it in production environments without proper testing and security considerations. I cannot be held responsible for any damage or issues caused by using this script. Use it at your own risk.
- If you would like to see improvements or have any suggestions, please feel free to contribute to the repository.

HOW TO USE IN YOUR OWN ENVIRONMENT:
- Make sure you have all the prerequisities to import Exchange Management Shell
- Make sure you have the necessary permissions to manage objects
- To make the script work properly, modify $pageSize and $yourdomain variables only

.NOTES
v20250418 - Initial version
v20250516 - Added ConvertToSharedMailbox and ConvertToUserMailbox functionality, also added for updating mailbox properties and creating new mailboxes (shared/user). Added disable archive functionality.
v20250517 - Added error handling for mailbox updates and improved HTML styling.
v20250520 - Added proxy address handling for email addresses in the UpdateMailbox function. Email address separator is now a semicolon.
v20250526 - A lot of improvements. Should be good for beta testing now.
v20250611 - Disable Archive button is moved to a good position.
v20250704 - Added OwnerN functionality to set the OwnerN attribute for specific accounts.
v20260402 - Added $yourdomain variable for easier configuration of email address generation in Enable-RemoteMailbox and UpdateMailbox functions. Updated HTML placeholders to use this variable as well.
#>

# Import Exchange Management Shell
Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction SilentlyContinue

# Use TLS 1.2 if required for any secure connections
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Create and start an HttpListener listening on http://localhost:8080/
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:8080/")
$listener.Start()
Write-Host "Exchange Control Panel running on http://localhost:8080/"

# Define a page size for paging
$pageSize = 50

# Define your domain for email address generation (used in Enable-RemoteMailbox and UpdateMailbox functions). Use only the domain part (without @).
$yourdomain = "domain"

# Template function to wrap content in common HTML with styling
function Get-HTMLPage {
    param(
        [string]$Title,
        [string]$Content
    )
    $htmlTemplate = @"
<html>
<head>
  <meta charset='utf-8'>
  <title>$Title</title>
  <style>
    body { font-family: Arial, sans-serif; background-color: #f4f4f4; color: #333; margin: 0; padding: 20px; }
    .container { max-width: 1024px; margin: 0 auto; background-color: #fff; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }
    h1 { color: #444; }
    nav a { margin-right: 10px; color: #007acc; text-decoration: none; }
    nav a:hover { text-decoration: underline; }
    table { border-collapse: collapse; width: 100%; background-color: #fff; }
    table, th, td { border: 1px solid #ccc; }
    th, td { padding: 10px; text-align: left; }
    tr:nth-child(even) { background-color: #f9f9f9; }
    .button { background-color: #007acc; color: #fff; border: none; padding: 10px 15px; text-decoration: none; cursor: pointer; border-radius: 3px; }
    .button:hover { background-color: #005f9b; }
    .button + .button { margin-left: 10px; }
    .nav { margin-bottom: 20px; }
    input[type='text'] { width: 500px; padding: 5px; font-size: 14px; border: 1px solid #ccc; border-radius: 3px; }
  </style>
</head>
<body>
  <div class='container'>
    $Content
  </div>
</body>
</html>
"@
    return $htmlTemplate
}

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $request = $context.Request
    $response = $context.Response

    try {
        $path = $request.Url.AbsolutePath

        switch ($path) {
            "/" {
                # Landing page
                $content = @"
<h1>Exchange Control Panel</h1>
<nav class='nav'>
  <a href='/ListMailboxes'>List All Mailboxes</a>|
  <a href='/ListSharedMailboxes'>List All Shared Mailboxes</a>|
  <a href='/ListDistributionGroups'>List All Distribution Groups</a>|
  <a href='/ListRoomEquipmentMailboxes'>List All Room/Equipment Mailboxes</a>
</nav>
<hr>
<h2>Search Mailbox</h2>
<form action='/SearchMailbox' method='get'>
  <input type='text' name='q' placeholder='Enter mailbox name'>
  <input type='submit' class='button' value='Search'>
</form>
<hr>
<h2>Search Distribution Group</h2>
<form action='/SearchDistributionGroup' method='get'>
  <input type='text' name='q' placeholder='Enter distribution group name'>
  <input type='submit' class='button' value='Search'>
</form>
<hr>
<h2>Search Room/Equipment Mailbox</h2>
<form action='/SearchRoomEquipmentMailbox' method='get'>
  <input type='text' name='q' placeholder='Enter room/equipment mailbox name'>
  <input type='submit' class='button' value='Search'>
</form>
<hr>
<h2>Enable Remote Mailbox</h2>
<form action='/EnableRemoteMailbox' method='post'>
  <p>
    <label>Alias:</label>
    <input type='text' name='alias' placeholder='Enter alias for the mailbox'>
    <input type='submit' class='button' value='Create'>
  </p>
</form>
<hr>
<h2>Set OwnerN Attribute</h2>
<form action='/SetOwnerN' method='post'>
  <p>
    <label>Account (UPN or sAMAccountName):</label>
    <input type='text' name='account' placeholder="user@$($yourdomain).com or SamAccountName">
  </p>
  <p>
    <label>OwnerN (UPN):</label>
    <input type='text' name='ownern' placeholder="owner@$($yourdomain).com">
  </p>
  <input type='submit' class='button' value='Set OwnerN'>
</form>
<hr>
"@
                $html = Get-HTMLPage -Title "Exchange Control Panel" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ListMailboxes" {
                # Paging: default to page 1 if not specified.
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                # Retrieve ALL remote mailboxes in the background
                $allMailboxes = Get-RemoteMailbox -SortBy Name -ResultSize Unlimited | Select-Object Name, UserPrincipalName, PrimarySmtpAddress, RecipientType, RecipientTypeDetails | Sort-Object Name
                $totalItems = $allMailboxes.Count
                # Select only the current page items
                $pagedMailboxes = $allMailboxes | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                $content = "<h1>All Mailboxes (Page $page)</h1>"
                $content += "<table><tr><th>Name</th><th>UserPrincipalName</th><th>Primary Email Address</th><th>RecipientTypeDetails</th></tr>"
                foreach ($mb in $pagedMailboxes) {
                    $content += "<tr><td><a href='/ViewMailbox?upn=$($mb.UserPrincipalName)'>$($mb.Name)</a></td><td>$($mb.UserPrincipalName)</td><td>$($mb.PrimarySmtpAddress)</td><td>$($mb.RecipientTypeDetails)</td></tr>"
                }
                $content += "</table>"

                # Paging navigation links
                if ($page -gt 1) {
                    $prevPage = $page - 1
                    $content += "<p><a class='button' href='/ListMailboxes?page=$prevPage'>Previous</a></p>"
                }
                if (($page * $pageSize) -lt $totalItems) {
                    $nextPage = $page + 1
                    $content += "<p><a class='button' href='/ListMailboxes?page=$nextPage'>Next</a></p>"
                }
                $content += "<p><a class='button' href='/'>Back to Home</a></p>"

                $html = Get-HTMLPage -Title "All Mailboxes" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ListSharedMailboxes" {
                # Paging: default to page 1 if not specified.
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                try {
                    # Retrieve all shared mailboxes
                    $allSharedMailboxes = Get-RemoteMailbox -Filter "RecipientTypeDetails -eq 'RemoteSharedMailbox'" -ResultSize Unlimited -SortBy Name | Select-Object Name, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails
                    $totalItems = $allSharedMailboxes.Count

                    # Select only the current page items
                    $pagedSharedMailboxes = $allSharedMailboxes | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                    $content = "<h1>All Shared Mailboxes (Page $page)</h1>"
                    $content += "<table><tr><th>Name</th><th>UserPrincipalName</th><th>Primary Email Address</th><th>RecipientTypeDetails</th></tr>"
                    foreach ($mb in $pagedSharedMailboxes) {
                        $content += "<tr><td><a href='/ViewMailbox?upn=$($mb.UserPrincipalName)'>$($mb.Name)</a></td><td>$($mb.UserPrincipalName)</td><td>$($mb.PrimarySmtpAddress)</td><td>$($mb.RecipientTypeDetails)</td></tr>"
                    }
                    $content += "</table>"

                    # Paging navigation links
                    if ($page -gt 1) {
                        $prevPage = $page - 1
                        $content += "<p><a class='button' href='/ListSharedMailboxes?page=$prevPage'>Previous</a></p>"
                    }
                    if (($page * $pageSize) -lt $totalItems) {
                        $nextPage = $page + 1
                        $content += "<p><a class='button' href='/ListSharedMailboxes?page=$nextPage'>Next</a></p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                catch {
                    $content = "<p>Error retrieving shared mailboxes: $($_.Exception.Message)</p>"
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "All Shared Mailboxes" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ListDistributionGroups" {
                # Determine current page (default to page 1 if none provided)
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                # Retrieve ALL Distribution Groups, sorted by DisplayName
                $allDGs = Get-DistributionGroup -SortBy DisplayName | Select-Object Name, Identity
                $totalItems = $allDGs.Count

                # Select only the current page items (display 50 per page)
                $pagedDGs = $allDGs | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                $content = "<h1>Distribution Groups (Page $page)</h1>"
                $content += "<table><tr><th>Name</th><th>Identity</th></tr>"
                foreach ($dg in $pagedDGs) {
                    $encodedId = [System.Web.HttpUtility]::UrlEncode($dg.Identity)
                    $content += "<tr><td><a href='/ViewDistributionGroup?identity=$encodedId'>$($dg.Name)</a></td><td>$($dg.Identity)</td></tr>"
                }
                $content += "</table>"

                # Add paging navigation links (Previous/Next)
                if ($page -gt 1) {
                    $prevPage = $page - 1
                    $content += "<p><a class='button' href='/ListDistributionGroups?page=$prevPage'>Previous</a></p>"
                }
                if (($page * $pageSize) -lt $totalItems) {
                    $nextPage = $page + 1
                    $content += "<p><a class='button' href='/ListDistributionGroups?page=$nextPage'>Next</a></p>"
                }
                $content += "<p><a class='button' href='/'>Back to Home</a></p>"

                $html = Get-HTMLPage -Title "Distribution Groups" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ListRoomEquipmentMailboxes" {
                # Define paging size
                $pageSize = 50

                # Determine current page (default to page 1 if none provided)
                $page = [int]$request.QueryString["page"]
                if (-not $page -or $page -lt 1) { $page = 1 }

                try {
                    # Retrieve all room and equipment mailboxes
                    $AllRoomEquipmentMBX = Get-RemoteMailbox -ResultSize Unlimited -Filter "RecipientTypeDetails -eq 'RemoteRoomMailbox' -or RecipientTypeDetails -eq 'RemoteEquipmentMailbox'"

                    if (-not $AllRoomEquipmentMBX -or $AllRoomEquipmentMBX.Count -eq 0) {
                        $content = "<p>No room/equipment mailboxes found.</p>"
                        $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                        $html = Get-HTMLPage -Title "Room/Equipment Mailboxes" -Content $content
                        $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                        $response.ContentType = "text/html"
                        $response.ContentLength64 = $buffer.Length
                        $response.OutputStream.Write($buffer, 0, $buffer.Length)
                        $response.OutputStream.Close()
                        return
                    }

                    $totalItems = $AllRoomEquipmentMBX.Count

                    # Select only the current page items (display 50 per page)
                    $pagedREMs = $AllRoomEquipmentMBX | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                    $content = "<h1>Room and Equipment Mailboxes (Page $page)</h1>"
                    $content += "<table><tr><th>DisplayName</th><th>Type</th><th>PrimarySMTPAddress</th></tr>"
                    foreach ($rem in $pagedREMs) {
                        $content += "<tr><td><a href='/ViewRoomEquipmentMailbox?identity=$($rem.Identity)'>$($rem.DisplayName)</a></td><td>$($rem.RecipientTypeDetails)</td><td>$($rem.PrimarySMTPAddress)</td></tr>"
                    }
                    $content += "</table>"

                    # Add paging navigation links (Previous/Next)
                    if ($page -gt 1) {
                        $prevPage = $page - 1
                        $content += "<p><a class='button' href='/ListRoomEquipmentMailboxes?page=$prevPage'>Previous</a></p>"
                    }
                    if (($page * $pageSize) -lt $totalItems) {
                        $nextPage = $page + 1
                        $content += "<p><a class='button' href='/ListRoomEquipmentMailboxes?page=$nextPage'>Next</a></p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                catch {
                    $content = "<p>Error retrieving room/equipment mailboxes: $($_.Exception.Message)</p>"
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "Room/Equipment Mailboxes" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ViewRoomEquipmentMailbox" {
                # View details for a Room/Equipment Mailbox
                $identity = $request.QueryString["identity"]
                $content = ""

                if (-not $identity) {
                    $content = "<p>No room/equipment mailbox specified. <br><a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $RoomEquipmentMBX = Get-RemoteMailbox -Identity $decodedId -ErrorAction SilentlyContinue
                    if ($RoomEquipmentMBX) {
                        $content = "<h1>$($RoomEquipmentMBX.Name)</h1>"
                        $content += "<p><strong>Identity:</strong> $($RoomEquipmentMBX.Identity)</p>"
                        $content += "<p><strong>Display Name:</strong> $($RoomEquipmentMBX.DisplayName)</p>"
                        $content += "<p><strong>Primary SMTP Address:</strong> $($RoomEquipmentMBX.PrimarySmtpAddress)</p>"
                        $content += "<p><strong>Alias:</strong> $($RoomEquipmentMBX.Alias)</p>"
                        $content += "<p><strong>Organizational Unit:</strong> $($RoomEquipmentMBX.OnPremisesOrganizationalUnit)</p>"
                        $content += "<p><strong>Recipient Type Details:</strong> $($RoomEquipmentMBX.RecipientTypeDetails)</p>"

                        # Add Disable, Delete, and Back buttons.
                        $content += "<p>
                        <a class='button' href='/ConfirmDisableRoomEquipmentMailbox?identity=$identity'>Disable Room/Equipment Mailbox</a>
                        <a class='button' href='/ConfirmDeleteRoomEquipmentMailbox?identity=$identity'>Delete Room/Equipment Mailbox</a>
                        </p>
                        <p><a class='button' href='/'>Back to Home</a></p>"
                    }

                    else {
                        $content = "<p>Room/Equipment mailbox not found. <a class='button' href='/'>Back to Home</a></p>"
                    }
                }

                $html = Get-HTMLPage -Title "Room/Equipment Mailbox Details" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConfirmDisableRoomEquipmentMailbox" {
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No room/equipment mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $content = "<h2>Confirm Disable Room/Equipment Mailbox</h2>"
                    $content += "<p>Are you sure you want to disable the room/equipment mailbox <strong>$decodedId</strong>?<br>"
                    $content += "This will remove Exchange properties from the object and cannot be undone easily.</p>"
                    $content += "<a class='button' href='/DisableRoomEquipmentMailbox?identity=$identity'>Yes, Disable</a> "
                    $content += "<a class='button' href='/ViewRoomEquipmentMailbox?identity=$identity'>Cancel</a>"
                }
                $html = Get-HTMLPage -Title "Confirm Disable Room/Equipment Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DisableRoomEquipmentMailbox" {
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No room/equipment mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    try {
                        Disable-RemoteMailbox -Identity $decodedId -Confirm:$false
                        $content = "<p>Room/Equipment mailbox <strong>$decodedId</strong> has been successfully disabled.</p>"
                    }
                    catch {
                        $content = "<p>Error disabling room/equipment mailbox <strong>$decodedId</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Disable Room/Equipment Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConfirmDeleteRoomEquipmentMailbox" {
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No room/equipment mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $content = "<h2>Confirm Delete Room/Equipment Mailbox</h2>"
                    $content += "<p>Are you sure you want to delete the room/equipment mailbox <strong>$decodedId</strong>?<br>"
                    $content += "This action will permanently remove the mailbox and its associated Active Directory object. This cannot be undone.</p>"
                    $content += "<a class='button' href='/DeleteRoomEquipmentMailbox?identity=$identity'>Yes, Delete</a> "
                    $content += "<a class='button' href='/ViewRoomEquipmentMailbox?identity=$identity'>Cancel</a>"
                }
                $html = Get-HTMLPage -Title "Confirm Delete Room/Equipment Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DeleteRoomEquipmentMailbox" {
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No room/equipment mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    try {
                        Remove-RemoteMailbox -Identity $decodedId -Confirm:$false
                        $content = "<p>Room/Equipment mailbox <strong>$decodedId</strong> has been successfully deleted.</p>"
                    }
                    catch {
                        $content = "<p>Error deleting room/equipment mailbox <strong>$decodedId</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Delete Room/Equipment Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/SearchMailbox" {
                # Determine page number; default to 1 if not present.
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                $query = $request.QueryString["q"]
                if ([string]::IsNullOrEmpty($query)) {
                    $content = "<p>No search term provided. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    # Retrieve all matching remote mailboxes (background query)
                    $allMailboxes = Get-RemoteMailbox -Anr "$query" | Sort-Object Name
                    $totalItems = $allMailboxes.Count
                    $pagedMailboxes = $allMailboxes | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                    $content = "<h1>Search Results for '$query' (Page $page)</h1>"
                    $content += "<table><tr><th>Name</th><th>UserPrincipalName</th></tr>"
                    foreach ($mb in $pagedMailboxes) {
                        $content += "<tr><td><a href='/ViewMailbox?upn=$($mb.UserPrincipalName)'>$($mb.Name)</a></td><td>$($mb.UserPrincipalName)</td></tr>"
                    }
                    $content += "</table>"
                    if ($page -gt 1) {
                        $prevPage = $page - 1
                        $content += "<p><a class='button' href='/SearchMailbox?q=$query&page=$prevPage'>Previous</a></p>"
                    }
                    if (($page * $pageSize) -lt $totalItems) {
                        $nextPage = $page + 1
                        $content += "<p><a class='button' href='/SearchMailbox?q=$query&page=$nextPage'>Next</a></p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Search Results" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ViewMailbox" {
                # View individual mailbox details – expects query parameter "upn"
                $upn = $request.QueryString["upn"]
                $content = ""

                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $mb = Get-RemoteMailbox -Identity $upn
                    if ($mb) {
                        $name = $mb.Name
                        $displayName = $mb.DisplayName
                        $alias = $mb.Alias
                        $primarySMTPAddress = $mb.PrimarySmtpAddress
                        $emailAddresses = ($mb.EmailAddresses -join "; ")
                        $userPrincipalName = $mb.UserPrincipalName
                        $remoteRoutingAddress = $mb.RemoteRoutingAddress
                        $mailboxType = $mb.RecipientTypeDetails
                        $isDisabled = $mb.AccountDisabled
                        $archiveName = $mb.ArchiveName
                        $archiveStatus = $mb.ArchiveStatus
                        $archiveGuid = $mb.ArchiveGuid

                        $actionButtons = @"
    <input type='submit' class='button' value='Update Mailbox'>
    <a class='button' href='/ConfirmDisableMailbox?upn=$upn'>Disable Mailbox</a>
    <a class='button' href='/ConvertToSharedMailbox?upn=$upn'>Convert to Shared</a>
    <a class='button' href='/ConvertToUserMailbox?upn=$upn'>Convert to User</a>
"@
                        if ($archiveStatus -and $archiveStatus -ne "None") {
                            $actionButtons += "<a class='button' href='/DisableArchive?upn=$upn'>Disable Archive</a>"
                        }

                        $content = @"
<h1>$name mailbox details:</h1>
<form method='post' action='/UpdateMailbox'>
  <input type='hidden' name='upn' value='$upn'>
  <p>
    <label><strong>Name:</strong></label>
    $name
  </p>
  <p>
    <label><strong>Display Name:</strong></label>
    $displayName
  </p>
  <p>
    <label><strong>Alias:</strong></label>
    $alias
  </p>
  <p>
    <label><strong>User Principal Name:</strong></label>
    $userPrincipalName
  </p>
  <p>
    <label><strong>Primary Email Address:</strong></label>
    <input type='text' name='PrimarySMTPAddress' value='$primarySMTPAddress'>
  </p>
  <p>
    <label><strong>Email Addresses:</strong></label>
    <input type='text' name='EmailAddresses' value='$emailAddresses'>
  </p>
  <p>
    <label><strong>RemoteRoutingAddress:</strong></label>
    $remoteRoutingAddress
  </p>
  <p>
    <label><strong>Mailbox type:</strong></label>
    $mailboxType
  </p>
    <p>
    <label><strong>Account disabled:</strong></label>
    $isDisabled
  </p>
  <p>
    <label><strong>Archive Name:</strong></label>
    $archiveName
  </p>
  <p>
    <label><strong>Archive Status:</strong></label>
    $archiveStatus
  </p>
  <p>
    <label><strong>Archive GUID:</strong></label>
    $archiveGuid
  </p>
  <p>
    $actionButtons
</p>
    <p>
    <a class='button' href='/'>Back to Home</a>
    </p>
    </form>
"@

                    }
                    else {
                        $content = "<p>Mailbox not found. <a class='button' href='/'>Back to Home</a></p>"
                    }
                }
                $html = Get-HTMLPage -Title "Mailbox Details - $upn" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/UpdateMailbox" {
                if ($request.HttpMethod -eq "POST") {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                    $data = $reader.ReadToEnd()
                    $reader.Close()
                    $parsed = [System.Web.HttpUtility]::ParseQueryString($data)
                    $upn = $parsed["upn"]
                    $newPrimarySMTPAddress = $parsed["PrimarySMTPAddress"]
                    $newEmailAddresses = $parsed["EmailAddresses"]
                    $newUPN = $parsed["UserPrincipalName"]
                    $newRemoteRoutingAddress = $parsed["RemoteRoutingAddress"]

                    try {
                        # Retrieve current mailbox properties
                        $currentMailbox = Get-RemoteMailbox -Identity $upn

                        # Prepare update parameters dynamically
                        $updateParams = @{}

                        # EmailAddresses
                        if ($newEmailAddresses -and $newEmailAddresses -ne ($currentMailbox.EmailAddresses -join "; ")) {
                            $newEmailAddressesArray = $newEmailAddresses -split ";" | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                            $currentEmailAddressesArray = $currentMailbox.EmailAddresses | ForEach-Object { "$_" } | ForEach-Object { $_.Trim() }
                            if ($newEmailAddressesArray -and -not (@($newEmailAddressesArray | Sort-Object) -ceq @($currentEmailAddressesArray | Sort-Object))) {
                                $proxyAddresses = $newEmailAddressesArray | ForEach-Object { [Microsoft.Exchange.Data.ProxyAddress]::Parse($_) }
                                $updateParams["EmailAddresses"] = $proxyAddresses
                            }
                        }

                        # UserPrincipalName
                        if ($newUPN -and $newUPN -ne $currentMailbox.UserPrincipalName) {
                            $updateParams["UserPrincipalName"] = $newUPN
                        }

                        # PrimarySMTPAddress and RemoteRoutingAddress logic
                        $primarySMTPChanged = $false
                        if ($newPrimarySMTPAddress -and $newPrimarySMTPAddress -ne $currentMailbox.PrimarySmtpAddress) {
                            $primarySMTPChanged = $true
                            # Extract alias from newPrimarySMTPAddress (before @)
                            $aliasPart = $newPrimarySMTPAddress.Split("@")[0]
                            $autoRemoteRoutingAddress = "$aliasPart@$($yourdomain).mail.onmicrosoft.com"
                            $updateParams["PrimarySmtpAddress"] = $newPrimarySMTPAddress
                            $updateParams["RemoteRoutingAddress"] = $autoRemoteRoutingAddress
                        }
                        elseif ($newRemoteRoutingAddress -and $newRemoteRoutingAddress -ne "$($currentMailbox.RemoteRoutingAddress)") {
                            $updateParams["RemoteRoutingAddress"] = $newRemoteRoutingAddress
                        }

                        # Apply updates only if there are changes
                        if ($updateParams.Count -gt 0) {
                            Set-RemoteMailbox -Identity $upn @updateParams -Confirm:$false
                            $result = "Successfully updated mailbox $upn."
                        }
                        else {
                            $result = "No changes detected for mailbox $upn."
                        }
                    }
                    catch {
                        $result = "Error updating mailbox: $($_.Exception.Message)"
                    }

                    $content = "<p>$result</p><p><a class='button' href='/ViewMailbox?upn=$upn'>Return to Mailbox Details</a></p><p><a class='button' href='/'>Back to Home</a></p>"
                    $html = Get-HTMLPage -Title "Update Result" -Content $content
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = "text/html"
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                }
                else {
                    $response.StatusCode = 405
                    $response.StatusDescription = "Method Not Allowed"
                    $response.OutputStream.Close()
                }
            }

            "/ConfirmDisableMailbox" {
                $upn = $request.QueryString["upn"]
                $content = ""
                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $content = "<h2>Confirm Disable Mailbox</h2>"
                    $content += "<p>Are you sure you want to disable the mailbox <strong>$upn</strong>?<br>"
                    $content += "This will remove Exchange properties from the user object and cannot be undone easily.</p>"
                    $content += "<a class='button' href='/DisableMailbox?upn=$upn'>Yes, Disable Mailbox</a> "
                    $content += "<a class='button' href='/ViewMailbox?upn=$upn'>Cancel</a>"
                }
                $html = Get-HTMLPage -Title "Confirm Disable Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DisableMailbox" {
                # Disable the remote mailbox – expects query parameter "upn"
                $upn = $request.QueryString["upn"]
                $content = ""

                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    try {
                        # Disable the remote mailbox without asking for confirmation.
                        Disable-RemoteMailbox -Identity $upn -Confirm:$false
                        $content = "<p>Mailbox <strong>$upn</strong> has been successfully disabled.</p>"
                    }
                    catch {
                        $content = "<p>Error disabling mailbox <strong>$upn</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Disable Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConvertToSharedMailbox" {
                # Convert the mailbox to a shared mailbox – expects query parameter "upn"
                $upn = $request.QueryString["upn"]
                $content = ""

                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    try {
                        # Convert the mailbox to a shared mailbox
                        Set-RemoteMailbox -Identity $upn -Type Shared
                        $content = "<p>Mailbox <strong>$upn</strong> has been successfully converted to a shared mailbox.</p>"
                    }
                    catch {
                        $content = "<p>Error converting mailbox <strong>$upn</strong> to shared: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/ViewMailbox?upn=$upn'>Return to Mailbox Details</a></p>"
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "Convert to Shared Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConvertToUserMailbox" {
                # Convert the mailbox to a user mailbox – expects query parameter "upn"
                $upn = $request.QueryString["upn"]
                $content = ""

                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    try {
                        # Convert the mailbox to a user mailbox
                        Set-RemoteMailbox -Identity $upn -Type Regular
                        $content = "<p>Mailbox <strong>$upn</strong> has been successfully converted to a user mailbox.</p>"
                    }
                    catch {
                        $content = "<p>Error converting mailbox <strong>$upn</strong> to user: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/ViewMailbox?upn=$upn'>Return to Mailbox Details</a></p>"
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "Convert to User Mailbox" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DisableArchive" {
                $upn = $request.QueryString["upn"]
                $content = ""

                if (-not $upn) {
                    $content = "<p>No mailbox specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    try {
                        # Disable the archive for the mailbox
                        Disable-Mailbox -Identity $upn -Archive -Confirm:$false
                        $content = "<p>Archive for mailbox <strong>$upn</strong> has been successfully disabled.</p>"
                    }
                    catch {
                        $content = "<p>Error disabling archive for mailbox <strong>$upn</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/ViewMailbox?upn=$upn'>Return to Mailbox Details</a></p>"
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "Disable Archive" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/EnableRemoteMailbox" {
                if ($request.HttpMethod -eq "POST") {
                    # Read the POST data
                    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                    $data = $reader.ReadToEnd()
                    $reader.Close()

                    # Parse the form data
                    $parsed = [System.Web.HttpUtility]::ParseQueryString($data)
                    $identity = $parsed["alias"]

                    if ([string]::IsNullOrEmpty($identity)) {
                        $content = "<p>AD username is required. <a class='button' href='/'>Back to Home</a></p>"
                    }
                    else {
                        try {
                            $remoteRoutingAddress = "$identity@$($yourdomain).mail.onmicrosoft.com"
                            $primarySMTPAddress = "$identity@$($yourdomain).com"

                            Enable-RemoteMailbox -Identity $identity -RemoteRoutingAddress $remoteRoutingAddress -PrimarySMTPAddress $primarySMTPAddress -Confirm:$false

                            $content = "<p>Remote mailbox for <strong>$identity</strong> has been successfully created as a <strong>User</strong> mailbox.</p>"
                        }
                        catch {
                            $content = "<p>Error creating remote mailbox for <strong>$identity</strong>: $($_.Exception.Message)</p>"
                        }
                        $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                    }

                    $html = Get-HTMLPage -Title "Enable Remote Mailbox" -Content $content
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = "text/html"
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                }
            }

            "/SetOwnerN" {
                if ($request.HttpMethod -eq "POST") {
                    $reader = New-Object System.IO.StreamReader($request.InputStream, $request.ContentEncoding)
                    $data = $reader.ReadToEnd()
                    $reader.Close()
                    $parsed = [System.Web.HttpUtility]::ParseQueryString($data)
                    $accountInput = $parsed["account"]
                    $ownerNInput = $parsed["ownern"]

                    # Define allowed OUs here
                    $allowedOUs = @(
                        "OU=__Generic,OU=2_R_EMEA,OU=1_Accounts,DC=niladv,DC=org",
                        "OU=ServiceAccounts,OU=2_Priv-Accounts,DC=niladv,DC=org",
                        "OU=__GENERIC,OU=3_R_APAC,OU=1_Accounts,DC=niladv,DC=org",
                        "OU=__GENERIC,OU=1_R_NCSA,OU=1_Accounts,DC=niladv,DC=org"
                    )

                    $result = ""
                    if ([string]::IsNullOrWhiteSpace($accountInput) -or [string]::IsNullOrWhiteSpace($ownerNInput)) {
                        $result = "Both fields are required."
                    }
                    else {
                        try {
                            # Try to resolve UPN to sAMAccountName or distinguishedName
                            $user = Get-ADUser -Filter { UserPrincipalName -eq $accountInput -or SamAccountName -eq $accountInput } -Properties DistinguishedName, UserPrincipalName, SamAccountName
                            if ($user) {
                                # Check if user is in any allowed OU
                                $isAllowed = $false
                                foreach ($ou in $allowedOUs) {
                                    if ($user.DistinguishedName -like "*$ou") {
                                        $isAllowed = $true
                                        break
                                    }
                                }
                                if ($isAllowed) {
                                    Set-ADUser -Identity $user.DistinguishedName -Replace @{ownerN = $ownerNInput }
                                    $result = "OwnerN set for $($user.UserPrincipalName) ($($user.SamAccountName))."
                                }
                                else {
                                    $result = "This account is not in an allowed OU."
                                }
                            }
                            else {
                                $result = "Account not found."
                            }
                        }
                        catch {
                            $result = "Error: $($_.Exception.Message)"
                        }
                    }
                    $content = "<p>$result</p><p><a class='button' href='/'>Back to Home</a></p>"
                    $html = Get-HTMLPage -Title "Set OwnerN Result" -Content $content
                    $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                    $response.ContentType = "text/html"
                    $response.ContentLength64 = $buffer.Length
                    $response.OutputStream.Write($buffer, 0, $buffer.Length)
                    $response.OutputStream.Close()
                }
            }

            "/SearchDistributionGroup" {
                # Determine current page number (default to 1 if not provided)
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                # Get the search query from the query string parameter "q"
                $query = $request.QueryString["q"]
                if ([string]::IsNullOrEmpty($query)) {
                    $content = "<p>No search term provided. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    # Retrieve ALL matching Distribution Groups in the background using a wildcard filter on Name
                    $allDGs = Get-DistributionGroup -Anr $query | Select-Object Name, Identity, WindowsEmailAddress
                    $totalItems = $allDGs.Count
                    # Select only the items for the current page (50 per page)
                    $pagedDGs = $allDGs | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                    $content = "<h1>Search Results for '$query' (Page $page)</h1>"
                    $content += "<table><tr><th>Name</th><th>Identity</th><th>Email address</th></tr>"
                    foreach ($dg in $pagedDGs) {
                        $encodedId = [System.Web.HttpUtility]::UrlEncode($dg.Identity)
                        $content += "<tr><td><a href='/ViewDistributionGroup?identity=$encodedId'>$($dg.Name)</a></td><td>$($dg.Identity)</td><td>$($dg.WindowsEmailAddress)</td></tr>"
                    }
                    $content += "</table>"

                    # If there is a previous page, include a Previous link
                    if ($page -gt 1) {
                        $prevPage = $page - 1
                        $content += "<p><a class='button' href='/SearchDistributionGroup?q=$query&page=$prevPage'>Previous</a></p>"
                    }
                    # If there are more items beyond the current page, include a Next link
                    if (($page * $pageSize) -lt $totalItems) {
                        $nextPage = $page + 1
                        $content += "<p><a class='button' href='/SearchDistributionGroup?q=$query&page=$nextPage'>Next</a></p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Search Distribution Groups" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ViewDistributionGroup" {
                # View details for a Distribution Group and its members, with buttons to disable and delete the group
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No distribution group specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $dg = Get-DistributionGroup -Identity $decodedId -SortBy DisplayName -ResultSize Unlimited
                    if ($dg) {
                        $content = "<h1>$($dg.Name)</h1>"
                        $content += "<p><strong>Identity:</strong> $($dg.Identity)</p>"
                        $content += "<p><strong>Organizational Unit:</strong> $($dg.OrganizationalUnit)</p>"
                        $content += "<p><strong>Description:</strong> $($dg.Notes)</p>"
                        $content += "<p><strong>Hidden from address list:</strong> $($dg.HiddenFromAddressListsEnabled)</p>"
                        $content += "<p><strong>Email address:</strong> $($dg.PrimarySmtpAddress)</p>"
                        $content += "<p><strong>Group type:</strong> $($dg.GroupType)</p>"
                        $content += "<p><strong>Owner approval is required to join the group?:</strong> $($dg.MemberJoinRestriction)</p>"
                        $content += "<p><strong>Is the group is open to leave?:</strong> $($dg.MemberDepartRestriction)</p>"
                        $content += "<p><strong>Managed By:</strong> $($dg.ManagedBy)</p>"

                        # Retrieve distribution group members
                        $members = Get-DistributionGroupMember -Identity $decodedId | Sort-Object Name
                        if ($members -and $members.Count -gt 0) {
                            $content += "<h2>Members</h2>"
                            $content += "<table><tr><th>Name</th><th>Email address</th></tr>"
                            foreach ($member in $members) {
                                $content += "<tr><td>$($member.Name)</td><td>$($member.PrimarySmtpAddress)</td></tr>"
                            }
                            $content += "</table>"
                        }
                        else {
                            $content += "<p><em>No members found in this group.</em></p>"
                        }
                        # Add Disable, Delete, and Back buttons.
                        $content += "<p>
                        <a class='button' href='/ConfirmDisableDistributionGroup?identity=$identity'>Disable Distribution Group</a>
                        <a class='button' href='/ConfirmDeleteDistributionGroup?identity=$identity'>Delete Distribution Group</a>
                        </p>
                        <p><a class='button' href='/'>Back to Home</a></p>"
                    }
                    else {
                        $content = "<p>Distribution group not found. <a class='button' href='/'>Back to Home</a></p>"
                    }
                }
                $html = Get-HTMLPage -Title "Distribution Group Details" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DisableDistributionGroup" {
                # Disable the distribution group – expects query parameter "identity"
                $identity = $request.QueryString["identity"]
                $content = ""

                if (-not $identity) {
                    $content = "<p>No distribution group specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    try {
                        # Actually disable the distribution group
                        Disable-DistributionGroup -Identity $decodedId -Confirm:$false
                        $content = "<p>Distribution group <strong>$decodedId</strong> has been successfully disabled.</p>"
                    }
                    catch {
                        $content = "<p>Error disabling distribution group <strong>$decodedId</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }

                $html = Get-HTMLPage -Title "Disable Distribution Group" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConfirmDisableDistributionGroup" {
                $identity = $request.QueryString["identity"]
                $content = ""
                if (-not $identity) {
                    $content = "<p>No distribution group specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $content = "<h2>Confirm Disable Distribution Group</h2>"
                    $content += "<p>Are you sure you want to disable the distribution group <strong>$decodedId</strong>?<br>"
                    $content += "This will remove Exchange properties from the user object and cannot be undone easily.</p>"
                    $content += "<a class='button' href='/DisableDistributionGroup?identity=$identity'>Yes, Disable Group</a> "
                    $content += "<a class='button' href='/ViewDistributionGroup?identity=$identity'>Cancel</a>"
                }
                $html = Get-HTMLPage -Title "Confirm Disable Distribution Group" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/DeleteDistributionGroup" {
                # Delete the distribution group using Remove-DistributionGroup
                $identity = $request.QueryString["identity"]
                $content = ""

                if (-not $identity) {
                    $content = "<p>No distribution group specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    try {
                        # Remove-DistributionGroup will delete the group including its AD object.
                        Remove-DistributionGroup -Identity $decodedId -Confirm:$false
                        $content = "<p>Distribution group <strong>$decodedId</strong> has been successfully deleted.</p>"
                    }
                    catch {
                        $content = "<p>Error deleting distribution group <strong>$decodedId</strong>: $($_.Exception.Message)</p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Delete Distribution Group" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/ConfirmDeleteDistributionGroup" {
                # Confirm deletion of a Distribution Group before actually removing it
                $identity = $request.QueryString["identity"]
                $content = ""

                if (-not $identity) {
                    $content = "<p>No distribution group specified. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    $decodedId = [System.Web.HttpUtility]::UrlDecode($identity)
                    $dg = Get-DistributionGroup -Identity $decodedId -ErrorAction SilentlyContinue
                    if ($dg) {
                        $content = "<h1>Confirm Delete Distribution Group</h1>"
                        $content += "<p>Are you sure you want to delete the distribution group <strong>$($dg.Name)</strong>? "
                        $content += "This action will permanently remove the group <em>and</em> delete the associated Active Directory object. "
                        $content += "This cannot be undone.</p>"
                        # Two buttons: one to confirm and one to cancel.
                        $content += "<p><a class='button' href='/DeleteDistributionGroup?identity=$identity'>Yes, Delete</a><br>"
                        $content += "<a class='button' href='/ViewDistributionGroup?identity=$identity'>Cancel</a></p>"
                    }
                    else {
                        $content = "<p>Distribution group not found. <a class='button' href='/'>Back to Home</a></p>"
                    }
                }
                $html = Get-HTMLPage -Title "Confirm Delete Distribution Group" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }

            "/SearchRoomEquipmentMailbox" {
                # Determine current page number (default to 1 if not provided)
                $page = [int]$request.QueryString["page"]
                $content = ""

                if (-not $page -or $page -lt 1) { $page = 1 }

                # Get the search query from the query string parameter "q"
                $query = $request.QueryString["q"]
                if ([string]::IsNullOrEmpty($query)) {
                    $content = "<p>No search term provided. <a class='button' href='/'>Back to Home</a></p>"
                }
                else {
                    # Retrieve ALL matching Room/Equipment Mailboxes using a wildcard filter on Name
                    $allMailboxes = Get-RemoteMailbox -Filter "Name -like '*$query*'" | Select-Object Name, Identity, PrimarySmtpAddress
                    $totalItems = $allMailboxes.Count
                    # Select only the items for the current page (50 per page)
                    $pagedMailboxes = $allMailboxes | Select-Object -Skip (($page - 1) * $pageSize) -First $pageSize

                    $content = "<h1>Search Results for '$query' (Page $page)</h1>"
                    $content += "<table><tr><th>Name</th><th>Identity</th><th>Email Address</th></tr>"
                    foreach ($mb in $pagedMailboxes) {
                        $encodedId = [System.Web.HttpUtility]::UrlEncode($mb.Identity)
                        $content += "<tr><td><a href='/ViewRoomEquipmentMailbox?identity=$encodedId'>$($mb.Name)</a></td><td>$($mb.Identity)</td><td>$($mb.PrimarySmtpAddress)</td></tr>"
                    }
                    $content += "</table>"

                    # If there is a previous page, include a Previous link
                    if ($page -gt 1) {
                        $prevPage = $page - 1
                        $content += "<p><a class='button' href='/SearchRoomEquipmentMailbox?q=$query&page=$prevPage'>Previous</a></p>"
                    }
                    # If there are more items beyond the current page, include a Next link
                    if (($page * $pageSize) -lt $totalItems) {
                        $nextPage = $page + 1
                        $content += "<p><a class='button' href='/SearchRoomEquipmentMailbox?q=$query&page=$nextPage'>Next</a></p>"
                    }
                    $content += "<p><a class='button' href='/'>Back to Home</a></p>"
                }
                $html = Get-HTMLPage -Title "Search Room/Equipment Mailboxes" -Content $content
                $buffer = [System.Text.Encoding]::UTF8.GetBytes($html)
                $response.ContentType = "text/html"
                $response.ContentLength64 = $buffer.Length
                $response.OutputStream.Write($buffer, 0, $buffer.Length)
                $response.OutputStream.Close()
            }
        }
    }
    catch {
        $errorMsg = "Server error: $($_.Exception.Message)"
        $buffer = [System.Text.Encoding]::UTF8.GetBytes($errorMsg)
        $response.ContentType = "text/plain"
        $response.ContentLength64 = $buffer.Length
        $response.OutputStream.Write($buffer, 0, $buffer.Length)
        $response.OutputStream.Close()
    }
}