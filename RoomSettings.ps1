## To enable scrips, Run powershell as admin then type Set-ExecutionPolicy RemoteSigned
#region    --- Transcript Open
$TranscriptTemp = [System.IO.Path]::GetTempFileName()               
Start-Transcript -path $TranscriptTemp | Out-Null
#endregion --- Transcript Open
#region    --- Main function header
$scriptFullname = $PSCommandPath ; if (!($scriptFullname)) {$scriptFullname =$MyInvocation.InvocationName }
#$scriptXML      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".xml"  ### replace .ps1 with .xml
#$scriptCSV      = $scriptFullname.Substring(0, $scriptFullname.LastIndexOf('.'))+ ".csv"  ### replace .ps1 with .csv
$scriptDir      = Split-Path -Path $scriptFullname -Parent
$scriptName     = Split-Path -Path $scriptFullname -Leaf
$scriptBase     = $scriptName.Substring(0, $scriptName.LastIndexOf('.'))
$psm1="$($scriptDir)\ITAutomator.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {Write-Host "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#$psm1="$($scriptDir)\ITAutomator M365.psm1";if ((Test-Path $psm1)) {Import-Module $psm1 -Force} else {Write-Host "Err 99: Couldn't find '$(Split-Path $psm1 -Leaf)'";Start-Sleep -Seconds 10;Exit(99)}
#endregion  --- Main function header
$props_ignore = @()
$props_ignore += "DisplayName"
$props_ignore += "UserPrincipalName"
$props_ignore += "Warnings"
$props_ignore += "LicenseInfo"
$props_ignore += "AccountEnabled"
#
$props_valid_intg = @()
$props_valid_strg = @()
$props_valid_bool = @()
$props_valid_perm = @()
#
$props_valid_intg += "Place_Capacity"
$props_valid_strg += "Place_City"
$props_valid_bool += "Place_MTREnabled"
$props_valid_strg += "Place_AudioDeviceName"
$props_valid_strg += "Place_VideoDeviceName"
#
$props_valid_bool += "Proc_AddOrganizerToSubject"
$props_valid_bool += "Proc_DeleteSubject"
$props_valid_bool += "Proc_DeleteAttachments"
$props_valid_bool += "Proc_DeleteComments"
$props_valid_strg += "Proc_AutomateProcessing"
$props_valid_bool += "Proc_AllowRecurringMeetings"
$props_valid_bool += "Proc_ProcessExternalMeetingMessages"
$props_valid_bool += "Proc_RemovePrivateProperty"
$props_valid_bool += "Proc_AddAdditionalResponse"
$props_valid_strg += "Proc_AdditionalResponse"
$props_valid_bool += "Proc_AllowConflicts"
$props_valid_intg += "Proc_BookingWindowInDays"
$props_valid_intg += "Proc_MaximumDurationInMinutes"
#
$props_valid_perm += "Perm_CalendarDefault"
Write-Host "-----------------------------------------------------------------------------"
Write-Host ("$scriptName        Computer:$env:computername User:$env:username PSver:"+($PSVersionTable.PSVersion.Major))
Write-Host ""
Write-Host "Room Mailbox Settings Reporter / Updater"
Write-Host ""
Write-Host "[R] Report room settings to a CSV file"
Write-Host "    Note: This report can be useful before changing room settings."
Write-Host ""
Write-Host "[U] Update room settings from a CSV file"
Write-Host "    Note: Move the CSV from the Report folder into the Update folder and adjust it to make changes."
Write-Host ""
Write-Host "Note: See README.md for details."
Write-Host "-----------------------------------------------------------------------------"
If (AskForChoice "Exclude License info from Reports? (Yes = faster, No=Requires Connect-MgGraph)" -Default 0) {
    $LicenseInfoInReports=$false
} else {
    $LicenseInfoInReports=$true
}   
#region Connections
if ($LicenseInfoInReports) { # Connect-MgGraph
    if (-not (Get-Command -Name Connect-MgGraph -ErrorAction SilentlyContinue)) {
        Write-Host "Connect-MgGraph is NOT available. You may need to install the PowerShell module."
        Write-Host 'Install-Module Microsoft.Graph'
        PressEnterToContinue
        exit
    }
    if (-not (Get-Command -Name Get-MgDomain -ErrorAction SilentlyContinue)) {
        Write-Host "Get-MgDomain is NOT available. You may need to install the PowerShell module."
        Write-Host 'Install-Module Microsoft.Graph.Identity.DirectoryManagement'
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        # Check if already connected to Microsoft Graph
        $mgContext = Get-MgContext
        if ($mgContext -and $mgContext.Account -and $mgContext.TenantId) {
            $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
            Write-Host "Connect-MgGraph is connected to Account $($mgContext.Account) Tenant Domain: " -NoNewline
            Write-Host $tenantDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-MgGraph..."
                Disconnect-MgGraph | Out-Null
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } else {
            Write-Host "Connect-MgGraph not connected. Connecting now..."
            PressEnterToContinue "Open a browser to an admin session on the desired tenant and press Enter."
            Connect-MgGraph -Scopes "User.ReadWrite.All", "Mail.ReadWrite", "Directory.ReadWrite.All" -NoWelcome
            # Confirm connection
            $mgContext = Get-MgContext
            if ($mgContext) {
                Write-Host "Now connected to Microsoft Graph as $($mgContext.Account)"
                $tenantDomain = (Get-MgDomain | Where-Object { $_.IsDefault }).Id
                Write-Host "Tenant Domain: $tenantDomain" -ForegroundColor Green
            } else {
                Write-Error "Connect-MgGraph: Failed"
            }
        }
    } # while true forever loop
    Write-Host
} # Connect-MgGraph
if ($true) { # Connect-ExchangeOnline
    # Check if Connect-ExchangeOnline is available
    if (-not (Get-Command Connect-ExchangeOnline -ErrorAction SilentlyContinue)) {
        Write-Host "Connect-ExchangeOnline is NOT available. You may need to install the PowerShell module."
        Write-Host 'Install-Module ExchangeOnlineManagement'
        PressEnterToContinue
        exit
    }
    # Check if we are already connected
    while ($true) {
        try {
            $orgConfig = Get-OrganizationConfig -ErrorAction Stop
            $connected_eol=$true
        }
        catch {
            $connected_eol=$false
        }
        if ($connected_eol)
        { # already connected
            # The Identity property typically shows your tenant's name or domain
            $tenantNameOrDomain = $orgConfig.Identity
            Write-Host "Connect-ExchangeOnline is connected to tenant: " -NoNewline
            Write-host $tenantNameOrDomain -ForegroundColor Green
            $response = AskForChoice "Choice: " -Choices @("&Use this connection","&Disconnect and try again","E&xit") -ReturnString
            # If the user types 'exit', break out of the loop
            if ($response -eq 'Disconnect and try again') {
                Write-Host "Disconnect-ExchangeOnline..."
                $null = Disconnect-ExchangeOnline -Confirm:$false
                PressEnterToContinue "Done. Press <Enter> to connect again."
                Continue # loop again
            }
            elseif ($response -eq 'exit') {
                return
            }
            else { # on to next step
                break
            }
        } # already connected
        else
        { # not connected
            Write-Host "Connect-ExchangeOnline is not connected."
            # check powershell version
            if ($PSVersionTable.PSVersion.Major -lt 7)
            { # PS 5
                Write-Host "We will try 'Connect-ExchangeOnline'. Use admin creds to authenticate."
                PressEnterToContinue
                Write-Host "Connect-ExchangeOnline (may be behind this window)... " -ForegroundColor Yellow
                Connect-ExchangeOnline -ShowBanner:$false
            } # PS 5
            else 
            { # PS 7
                $choice = AskForChoice "Connect method (PS7)" -Choices @("&Browser code","&Password","E&xit") -Default 0 -ReturnString
                if ($choice -eq "Exit") { exit }
                if ($choice -eq "Browser code")
                { # Browser code
                    Write-Host "1. Open a browser to an admin session on the desired tenant"
                    Write-Host "2. Copy the code below"
                    Write-Host "3. Click the link and paste the code to authenticate"
                    Write-Host "Connect-ExchangeOnline -Device ... " -ForegroundColor Yellow
                    Connect-ExchangeOnline -ShowBanner:$false -Device
                } # Browser code
                else { # Password
                    Write-Host "We will try 'Connect-ExchangeOnline' to authenticate. Use admin creds to authenticate."
                    PressEnterToContinue
                    Write-Host "Connect-ExchangeOnline (may be behind this window)... " -ForegroundColor Yellow
                    Connect-ExchangeOnline -ShowBanner:$false
                } # Password
            } # PS 7
            Write-Host "Done" -ForegroundColor Yellow
            Continue # loop again
        } # not connected
    } # while true forever loop
    Write-Host
} # Connect-ExchangeOnline
#endregion Connections
do
{
    $choice = AskForChoice "Choices" -Choices @("&Report RoomSettings","&Update RoomSettings","&Open this folder","E&xit") -Default 0 -ReturnString -ShowMenu
    if ($choice -eq "Exit") { Continue }
    if ($choice -eq "Open this folder") { Invoke-Item (Split-Path $scriptFullname -Parent); PressEnterToContinue "Folder opened in Explorer: $(Split-Path $scriptFullname -Parent)"; Continue }
    if ($choice -eq "Report RoomSettings") { # Report
        $rows=@()
        Write-Host "Get-Mailbox -RecipientTypeDetails RoomMailbox (list of Rooms) ... " -NoNewline
        # Get a list of room mailboxes in the tenant 
        $room_mbs = Get-Mailbox -RecipientTypeDetails RoomMailbox 
        Write-Host $room_mbs.count -ForegroundColor Green
        $i=0
        ForEach ($room_mb in $room_mbs)
        { # each mailbox
            Write-Host "$((++$i)) of $($room_mbs.Count): $($room_mb.DisplayName) <$($room_mb.UserPrincipalName)>"
            # Get calendar processing settings
            Write-Host " Get-CalendarProcessing.." -NoNewline
            $room_mb_proc = Get-CalendarProcessing -identity $room_mb.UserPrincipalName
            # Get place info (for capacity)
            Write-Host " Get-Place.." -NoNewline
            $room_mb_place = Get-Place -Identity $room_mb.UserPrincipalName
            # Shows the Default and Anonymous entries on the room’s Calenda
            Write-Host " Get-MailboxFolderPermission.." -NoNewline
            $room_mb_perm = Get-MailboxFolderPermission -Identity "$($room_mb.UserPrincipalName):\Calendar" | Where-Object { $_.User.UserType.Value -eq 'Default' } | Select-Object -First 1
            $room_mb_perm_str = if ($room_mb_perm) { $room_mb_perm.AccessRights -join "," } else { "None" }
            # License info
            if ($LicenseInfoInReports) {
                Write-Host " Get-MgUser.." -NoNewline
                $room_mb_user = Get-MgUser -UserId $room_mb.UserPrincipalName -Property "AssignedLicenses,DisplayName,AccountEnabled" -ErrorAction SilentlyContinue
                if ($room_mb_user) {
                    $assigned_licenses = $room_mb_user.AssignedLicenses | ForEach-Object { $_.SkuId }
                    $license_names = @()
                    foreach ($skuId in $assigned_licenses) {
                        $sku = Get-MgSubscribedSku | Where-Object { $_.SkuId -eq $skuId }
                        if ($sku) {
                            $license_names += $sku.SkuPartNumber
                        }
                    }
                    $room_mb_lic = ($license_names -join "; ")
                    if ($room_mb_lic -eq "") {
                        $room_mb_lic = "(None)"
                    }
                    $room_mb_userenabled = $room_mb_user.AccountEnabled
                } else {
                    $room_mb_lic = "(User not found)"
                }
            }
            else {
                $room_mb_lic = "(not checked)"
                $room_mb_userenabled = "(not checked)"
            }
            Write-Host " "
            #region Warnings
            $Warnings = @()
            if (-not $room_mb_proc) {
                $Warnings += "No calendar processing settings found"
            }
            if ($room_mb_proc.AutomateProcessing -ne "AutoAccept") {
                $Warnings += "AutomateProcessing should be AutoAccept"
            }
            if ($room_mb_proc.AllowRecurringMeetings -eq $false) {
                $Warnings += "Recurring meetings are NOT allowed"
            }
            if ($room_mb_proc.AllowConflicts -eq $true) {
                $Warnings += "AllowConflicts is TRUE (may lead to double-booking)"
            }
            if ($room_mb_proc.ProcessExternalMeetingMessages -eq $true) {
                $Warnings += "External meeting messages are processed (may allow spam bookings)"
            }
            if ($room_mb_proc.DeleteComments -eq $true) {
                $Warnings += "Comments (message body) is deleted - may remove 3rd party links"
            }
            if (($room_mb_proc.DeleteSubject -eq $true) -and ($room_mb_proc.AddOrganizerToSubject -eq $false)) {
                $Warnings += "Subjects are deleted - but organizer is NOT added to subject - may make it hard to identify meetings "
            }
            if ($room_mb_perm_str -notin ("AvailabilityOnly","LimitedDetails")) {
                $Warnings += "Calendar Default permission is '$room_mb_perm_str' (should be 'AvailabilityOnly' for free-busy or 'LimitedDetails' for subject-only)"
            }
            if ($room_mb_lic -eq "(None)") {
                $Warnings += "No license assigned"
            }
            if ($room_mb_userenabled -eq $false) {
                $Warnings += "Account is disabled"
            }
            if ($Warnings.count -gt 0) {
                Write-Host "  Warnings: " -NoNewline
                Write-Host ($Warnings -join "; ") -ForegroundColor Yellow
            }
            #endregion Warnings
            # Add row
            $row_obj=[pscustomobject][ordered]@{
                DisplayName                 = $room_mb.DisplayName
                UserPrincipalName           = $room_mb.UserPrincipalName
                Perm_CalendarDefault        = $room_mb_perm_str
                Place_Capacity                = $room_mb_place.Capacity
                Place_City                    = $room_mb_place.City
                Place_MTREnabled              = $room_mb_place.MTREnabled
                Place_AudioDeviceName         = $room_mb_place.AudioDeviceName
                Place_VideoDeviceName         = $room_mb_place.VideoDeviceName
                Proc_AddOrganizerToSubject          = $room_mb_proc.AddOrganizerToSubject
                Proc_DeleteSubject                  = $room_mb_proc.DeleteSubject
                Proc_DeleteAttachments              = $room_mb_proc.DeleteAttachments
                Proc_DeleteComments                 = $room_mb_proc.DeleteComments
                Proc_AutomateProcessing             = $room_mb_proc.AutomateProcessing
                Proc_AllowRecurringMeetings         = $room_mb_proc.AllowRecurringMeetings
                Proc_ProcessExternalMeetingMessages = $room_mb_proc.ProcessExternalMeetingMessages
                Proc_RemovePrivateProperty          = $room_mb_proc.RemovePrivateProperty
                Proc_AddAdditionalResponse          = $room_mb_proc.AddAdditionalResponse
                Proc_AdditionalResponse             = $room_mb_proc.AdditionalResponse
                Proc_AllowConflicts                 = $room_mb_proc.AllowConflicts
                Proc_BookingWindowInDays            = $room_mb_proc.BookingWindowInDays
                Proc_MaximumDurationInMinutes       = $room_mb_proc.MaximumDurationInMinutes
                AccountEnabled              = $room_mb_userenabled
                LicenseInfo                 = $room_mb_lic
                Warnings                    = $Warnings -join "; "
            }
            #### Add to results
            $rows += $row_obj
        } # each mailbox
        if ($rows.count -eq 0)
        { # no rows
            Write-Host "No room mailboxes found." -ForegroundColor Yellow
            PressEnterToContinue
        } # no rows
        else
        { # rows found
            # Export results
            $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $ReportTarget = "$(Split-Path $scriptFullname -Parent)\Reports\$($scriptBase)_report_$($date).csv"
            New-Item (Split-Path $ReportTarget -Parent) -ItemType Directory -Force | Out-Null # Make folder
            $rows | Export-Csv -Path $ReportTarget -NoTypeInformation -Encoding UTF8
            Write-Host "Report exported to: " -NoNewline
            Write-Host (Split-Path $ReportTarget -Leaf) -ForegroundColor Green
            if (AskForChoice "Open the report now?") {
                Invoke-Item $ReportTarget
            }
        } # rows found
    } # Report
    if ($choice -eq "Update RoomSettings") { # Update
        #region --- read / create CSV file of users
        $UpdateFolder = "$(Split-Path $scriptFullname -Parent)\Updates"
        New-Item $UpdateFolder -ItemType Directory -Force | Out-Null # Make folder
        # Search updates folder for most recent CSV file
        $UpdateFile = Get-ChildItem -Path $UpdateFolder -Filter "*.csv" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($UpdateFile) {
            $UpdateCSV = $UpdateFile.FullName
            Write-Host "Using most recent CSV file found in 'Updates' folder: " -NoNewline
            Write-Host (Split-Path $UpdateCSV -Leaf) -ForegroundColor Green
        } else {
            $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $UpdateCSV = "$($UpdateFolder)\$($scriptBase)_update_$($date).csv"
            Write-Host "No CSV files found in 'Updates' folder. Will use default file: " -NoNewline
            Write-Host (Split-Path $UpdateCSV -Leaf) -ForegroundColor Green
            ######### Template
            $row = "DisplayName,UserPrincipalName,"
            $row += ($props_valid_bool + $props_valid_strg + $props_valid_intg + $props_valid_perm) -join ","
            $row | Add-Content $UpdateCSV
            $row = "John Smith,john.smith@domain.com,FALSE,FALSE"
            $row | Add-Content $UpdateCSV
            ######### Template
            Write-Host "Template CSV created." -ForegroundColor Yellow
            PressEnterToContinue "Press <Enter> to open the CSV file for editing."
            Invoke-Item $UpdateCSV
            PressEnterToContinue "When done editing the CSV file, press <Enter> to continue."
        }
        ## ----------Fill $rows with contents of file
        $rows=@(import-csv $UpdateCSV)
        $rowscount = $rows.count
        Write-Host "CSV: $(Split-Path $UpdateCSV -leaf) ($($rowscount) entries)"
        $rows | Format-Table
        #endregion --- read / create CSV file of users
        $processed=0
        if ($true)
        { ## continue choices
            $choiceLoop=0
            $i=0        
            foreach ($row in $rows)
            { # each row
                $i++
                write-host "----- $i of $rowscount $($rows[0].PSObject.Properties | Select-Object -First 1 -ExpandProperty Value)"
                if ($choiceLoop -ne "Yes to All")
                {
                    $choiceLoop = AskForChoice "Process entry $($i) ?" -Choices @("&Yes","Yes to &All","&No","No and E&xit") -Default 1 -ReturnString
                }
                if (($choiceLoop -eq "Yes") -or ($choiceLoop -eq "Yes to All"))
                { # choiceloop
                    $processed++
                    #region    --------- Custom code for object $row
                    # Find mailbox
                    $room_mb = Get-Mailbox -Identity $row.UserPrincipalName -ErrorAction SilentlyContinue
                    if (-not $room_mb) {
                        Write-Host "Maibox with UserPrincipalName '$($row.UserPrincipalName)' not found. [And will be skipped]" -ForegroundColor Yellow
                        PressEnterToContinue
                        Continue
                    }
                    Write-Host "$($i): $($row.UserPrincipalName) [$($room_mb.DisplayName)]"
                    # Make sure we have a room mailbox
                    if ($room_mb.RecipientTypeDetails -ne "RoomMailbox") {
                        Write-Host "Mailbox with UserPrincipalName '$($row.UserPrincipalName)' is not a RoomMailbox (it's '$($room_mb.RecipientTypeDetails)'). [And will be skipped]" -ForegroundColor Yellow
                        PressEnterToContinue
                        Continue
                    }
                    # Get calendar processing settings
                    $room_mb_proc = Get-CalendarProcessing -identity $room_mb.UserPrincipalName
                    # Get place info (for capacity)
                    $room_mb_place = Get-Place -Identity $room_mb.UserPrincipalName
                    # Shows the Default and Anonymous entries on the room’s Calendar
                    $room_mb_perm = Get-MailboxFolderPermission -Identity "$($room_mb.UserPrincipalName):\Calendar" | Where-Object { $_.User.UserType.Value -eq 'Default' } | Select-Object -First 1
                    $room_mb_perm_str = if ($room_mb_perm) { $room_mb_perm.AccessRights -join "," } else { "None" }
                    # Loop through each property in the row
                    $row_props = $row.PSObject.Properties | Where-Object { $_.Name -notin $props_ignore }
                    foreach ($prop in $row_props)
                    { # row_props
                        $propName = $prop.Name
                        $propValue = $prop.Value
                        Write-Host "  $($propName): " -NoNewline
                        Write-Host "$($propValue) " -ForegroundColor Cyan -NoNewline
                        if ($propValue -in @("","<ignore>")) {
                            Write-Host "(OK: Ignored - blank value)" -ForegroundColor Magenta
                            Continue
                        } # empty value - ignore
                        if ($propName -like "*_*") {
                            $propertySource = $propName.Split("_")[0]
                            $propertyName = $propName.Split("_")[1]
                        }
                        else {
                            $propertySource = "Basic"
                            $propertyName = $propName
                        }
                        if ($propName -in ($props_valid_bool + $props_valid_strg + $props_valid_intg + $props_valid_perm))
                        { # valid properties
                            # get oldValue based on property source
                            if ($propertySource -eq "Proc")      { $oldValue = $room_mb_proc.$propertyName }
                            elseif ($propertySource -eq "Place") { $oldValue = $room_mb_place.$propertyName }
                            elseif ($propertySource -eq "Perm")  { $oldValue = $room_mb_perm_str }
                            else { Write-Host "Unknown property type: $($propertySource)" -ForegroundColor Red; Continue }
                            # get newvalue based on property type
                            if ($propName -in $props_valid_bool)
                            { # boolean properties
                                # Convert to boolean if needed
                                if ($propValue -in @("True","true","1")) { $newValue = $true }
                                elseif ($propValue -in @("False","false","0")) { $newValue = $false }
                                else { $newValue = $propValue } # leave as-is
                            } # boolean properties
                            elseif ($propName -in $props_valid_strg)
                            { # string properties
                                # Convert to permission array if needed
                                if ($propValue -as [string]) { $newValue = $propValue }
                                else { $newValue = $propValue } # leave as-is
                            } # string properties
                            elseif ($propName -in $props_valid_intg)
                            { # integer properties
                                # Convert to int if needed
                                if ($propValue -as [int]) { $newValue = $propValue }
                                else { $newValue = $propValue } # leave as-is
                            } # integer properties
                            elseif ($propName -in $props_valid_perm)
                            { # permission properties
                                $newValue = $propValue 
                            } # permission properties
                            else { Write-Host "Unknown property name: $($propertyName)" -ForegroundColor Red; Continue }
                            # Compare old and new values
                            if ($newValue -eq $oldValue) {
                                Write-Host "(OK: already set)" -ForegroundColor Green
                            } else {
                                Write-Host "(OK: changing from $($oldValue))" -ForegroundColor Yellow
                                # Update setting
                                try {   
                                    if ($propertySource -eq "Place") {
                                        Set-Place @params -ErrorAction Stop
                                    }
                                    elseif ($propertySource -eq "Proc") {
                                        Set-CalendarProcessing @params -ErrorAction Stop
                                    }
                                    elseif ($propertySource -eq "Perm") {
                                        # Update permission
                                        try {   
                                            Set-MailboxFolderPermission -Identity "$($room_mb.UserPrincipalName):\Calendar" -User Default -AccessRights $newValue -ErrorAction Stop
                                            Write-Host "  Updated." -ForegroundColor Green
                                        } catch {
                                            Write-Host "  Error updating permission: $_" -ForegroundColor Red
                                        }
                                    }
                                    Write-Host "  Updated." -ForegroundColor Green
                                } catch {
                                    Write-Host "  Error updating setting: $_" -ForegroundColor Red
                                }
                            }
                        } # valid properties
                        else {
                            Write-Host "(Unknown field - ignored)" -ForegroundColor Yellow
                        } # propName
                    } # row_props
                    #endregion --------- Custom code for object $row
                } # choiceloop
                if ($choiceLoop -eq "No")
                {
                    write-host ("Entry ($i) skipped.")
                }
                if ($choiceLoop -eq "No and Exit")
                {
                    write-host "Aborting."
                    Break
                }
            } # each row
        } ## continue choices
        Write-Host "------------------------------------------------------------------------------------"
        Write-Host "Done. $($processed) of $($rowscount) entries processed. Press [Enter] to exit."
        Write-Host "------------------------------------------------------------------------------------"
    } # Update
} until ($choice -eq "Exit") # loop until exit
#region    ---- Transcript Save
Stop-Transcript | Out-Null
$date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$TranscriptTarget = "$(Split-Path $scriptFullname -Parent)\Logs\$($scriptBase)_$($date)_log.txt"
New-Item (Split-Path $TranscriptTarget -Parent) -ItemType Directory -Force | Out-Null # Make Logs folder
Move-Item $TranscriptTemp $TranscriptTarget -Force
#endregion ---- Transcript Save