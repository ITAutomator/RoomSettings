
# RoomSettings.ps1

<img src=https://raw.githubusercontent.com/ITAutomator/Assets/main/RoomSettings/RoomSettings.png alt="screenshot" width="800"/>
 
A small interactive PowerShell utility to report and update Teams Room settings.  
See IT Automator script here: [Room Settings](https://github.com/ITAutomator/RoomSettings)

# Overview

Room settings can be configured to adjust the subject, message, and other details of an appointment to protect privacy and obscure information.  
However, most of these settings are not exposed in Microsoft's admin web sites. You will need to use PowerShell to adjust them.  
Change these settings carefully and according to the use case of your rooms.  
  
This script examines and adjusts room settings using these cmdlets.  
You can use the script to report settings across all your rooms (as a CSV file).  
You can then re-run the script in update mode, using the CSV file to provide the values needed.  
The update will only change things if required, so it's safe to run for rooms that are already configured.  
  
```powershell
Connect-ExchangeOnline
Get-CalendarProcessing
Set-CalendarProcessing
Get-MailboxFolderPermission
Set-MailboxFolderPermission

Connect-MgGraph
Get-Place
Set-Place
```
[Set-CalendarProcessing](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-calendarprocessing)  
[Set-MailboxFolderPermission](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxfolderpermission)  
[Set-Place](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-place)  
  
# Usage

Double-click `RoomSettings.cmd` to show the main menu.  
*Note: you will be prompted for M365 credentials*  
  
Choose `[R] Report RoomSettings` to create a CSV file for your rooms.  
Examine the settings, using the notes below to understand what they do.  
  
To update settings:  
Copy the CSV file from the `Reports` folder to the `Updates` folder.  
Edit the file as needed according to how you want your room settings.  
If you don't want a setting to be touched, you can delete the column. See **Columns** below for details.  
If you don't want a room to be touched, you can delete the row.  
  
Choose `[U] Update RoomSettings` to make your changes.  
The script automatically picks the most recent CSV file from the `Updates` folder 
If a setting is already correct, it will not be touched and reported as such.  
Rooms are processed one by one, pausing for input.  

If you are having trouble with **Join** button not showing for your meeting, see the **Troubleshooting / Notes** section below. 

# Settings 

## Microsoft details about these settings

Microsoft Source: [How to create and configure resource accounts for Teams Rooms and
panels](https://learn.microsoft.com/en-us/microsoftteams/rooms/create-resource-account?tabs=exchange-online%2Cgraph-powershell-password#configure-mailbox-properties)

### Exchange mailbox properties

Based on organizational requirements, you may wish to customize how the
resource account responds to and processes meeting invitations. Using
Exchange PowerShell, you can set many properties, review
the [Set-CalendarProcessing](https://learn.microsoft.com/en-us/powershell/module/exchange/mailboxes/set-calendarprocessing) cmdlet
for all available configurations.\
*The following are **Microsoft recommended for Teams Rooms (in bold)**:*

-   **AutomateProcessing: AutoAccept** \
    Meeting organizers receive the room reservation decision directly
    without human intervention.\
    *General Recommendation: Leave this alone.*

-   **AddOrganizerToSubject: \$false**\
    The meeting organizer isn\'t added to the subject of the meeting
    request on the resource account calendar.\
    *General Recommendation: Leave false. Change to true if
    DeleteSubject is true so that the room shows the organizer in the
    title.*

-   **AllowRecurringMeetings: \$true**\
    Recurring meetings are accepted.
    *General Recommendation: Leave true. Based on company policy.*

-   **DeleteAttachments: \$true**\
    Teams Rooms devices can\'t access meeting attachments, deleting
    attachments ensures they\'re not stored on the resource account
    calendar.
    *General Recommendation: Leave true.*

-   **DeleteComments: \$false**\
    Keep any text in the message body of incoming meeting requests which
    is required to create join buttons for [third-party
    meetings](https://learn.microsoft.com/en-us/microsoftteams/rooms/third-party-join) on
    a Teams Rooms device.
    *General Recommendation: Leave false. Warning: changing to true will break 3rd party meetings (see link above) like Zoom. Teams is not affected by this -- it uses other internal properties.*

-   **DeleteSubject: \$false**\
    Keep the subject of incoming meeting requests on the resource
    accounts calendar.
    *General Recommendation: Leave false. Change to true to hide subject (but then also set AddOrganizerToSubject to true)*

-   **ProcessExternalMeetingMessages: \$true**\
    Specifies whether to process meeting requests organized outside your Exchange environment. This option is required for meeting invites sent directly by an external organizer *as well as* external organized meetings *forwarded by* an internal user.\
    *General Recommendation: Leave true - unless there is concern about spam bookings from outside the company*

-   **RemovePrivateProperty: \$false**\
    Ensures the private flag that sent by the meeting organizer in the original meeting request remains as specified.

-   **AddAdditionalResponse: \$true**\
    The text specified by the AdditionalResponse parameter is added to meeting requests.

-   **AdditionalResponse: \"This is a Microsoft Teams Meeting room!\"**\
    The text to add to the meeting acceptance body. You can also format HTML content in the automatic reply if you wish.\
    *General Recommendation: Don't change these*

-   **MTREnabled: \$false**\
    True means\
    The room is marked as a Microsoft Teams Room. This typically means: The mailbox represents a physical space with a Teams Room device (like a Surface Hub or certified MTR console).\
    Teams and Outlook will treat the room as a "Teams Meeting Room," showing "Join" buttons on the device and enabling room-specific experiences.\
    This value may be set manually or via provisioning (e.g., when creating an MTR account following Microsoft's guidance).  

    False means\
    The room is a normal meeting room without any Teams Room device integration. It can still be scheduled for Teams meetings, but it won't behave as a dedicated MTR device\
    *General Recommendation: Don't change this. It seems to always be false for now even though the docs say it should be true*  

-   **Capacity: 12**\
    The number of people that the room can fit. Informational\
    *General Recommendation: Set to room capacity*

-   **Calendar Permission: AvailabilityOnly**\
    This is the calendar *default* permission (that is visible to all Org). More Info from Microsoft: [link](https://learn.microsoft.com/en-us/powershell/module/exchangepowershell/set-mailboxfolderpermission?view=exchange-ps#-accessrights)\
    *AvailabilityOnly* means show free-busy information only (not even subject is visible)\
    *LimitedDetails* means show subject but nothing else\
    *GeneralvRecommendation:*\
    *AvailabilityOnly* for most rooms.\
    *LimitedDetails* if *DeleteSubject* is true *and* there is an org-wide need to find the person who booked the room in case of conflicts. This would be the case if there were no known room personnel to check the calendar.\
    *LimitedDetails* if *DeleteSubject* is false for small orgs that are OK exposing the title of every booked meeting.  

# Reports

The CSV reports outputs the following information for all Room Mailboxes in the tenant (whether they are Teams Room licensed or not)\
The Updates CSV uses these same fields  

## Columns

| Column                                 | Value (sample)           | Description (*see info above) |
|--------                                |--------                  |-------------|
| DisplayName                            | Marketing Conference Rm  | (view-only info)             |
| UserPrincipalName                      | mkt_conf_room@domain.com | (required key field)         |
| Perm_CalendarDefault                   | AvailabilityOnly         | * AvailabilityOnly or LimitedDetails if Org wants to see room calendar |
| Proc_AddOrganizerToSubject             | TRUE                     | *TRUE Appends the organizer name to the subject       |
| Proc_DeleteSubject                     | TRUE                     | *TRUE Strips Subject (use with AddOrganizerToSubject) |
| Proc_DeleteAttachments                 | TRUE                     | *TRUE Strips Attachments                              |
| Proc_DeleteComments                    | FALSE                    | *TRUE Strips Body (caution: removes 3rd party links)  |
| Proc_ProcessExternalMeetingMessages    | TRUE                     | *FALSE: Organizer must be internal               |
| Proc_AutomateProcessing                | AutoAccept               | This should always be AutoAccept                 |
| Proc_AllowRecurringMeetings            | TRUE                     | FALSE: Do not allow recurring meetings           |
| Proc_RemovePrivateProperty             | FALSE                    | TRUE: Removes the property called Private |
| Proc_AddAdditionalResponse             | FALSE                    | TRUE: Add text from below field           |
| Proc_AdditionalResponse                |                          | Text response to append to booking response - generally blank |
| Proc_AllowConflicts                    | FALSE                    | TRUE: Allow overlapping meetings          |
| Proc_BookingWindowInDays               | 180                      | How far advanced booking can be in days   |
| Proc_MaximumDurationInMinutes          | 1440                     | How long a meeting can be in mins         |
| Place_Capacity                         | 8                        | The # of people the room can have  |
| Place_City                             | London                   |                                    |
| Place_MTREnabled                       | FALSE                    | (Generally false     )             |
| Place_AudioDeviceName                  |                          | Descriptive info - generally blank |
| Place_VideoDeviceName                  |                          | Descriptive info - generally blank |
| AccountEnabled                         | TRUE                     | (view-only info)                          |
| LicenseInfo                            | Teams Room Basic         | License assigned to room (if any) (view-only info) |
| PasswordExpiryDate                     |                          | (if the org expires passwords)                     |
| Warnings                               |                          | Any warnings detected by this code are shown here  |

## Warnings

Here are the various warnings that the code detects  
- AutomateProcessing should be AutoAccept  
- Recurring meetings are NOT allowed  
- AllowConflicts is TRUE (may lead to double-booking)  
- External meeting messages are processed (may allow spam bookings)  
- Comments (message body) is deleted - may remove 3rd party links  
- Subjects are deleted - but organizer is NOT added to subject - may make it hard to identify meetings   
- Calendar Default permission is '$room_mb_perm_str' (should be 'AvailabilityOnly' for free-busy or 'LimitedDetails' for subject-only)  
- No license assigned  
- Accounts with Teams room licenses should have a DisablePasswordExpiration policy (if the org expires passwords)  
- Account is disabled  

# Troubleshooting / Notes

## Deleting the Message Body

This will remove 3rd party links (Zoom). Teams is not affected by this. See here for more info: [Third-party meetings on Teams Rooms](https://learn.microsoft.com/en-us/microsoftteams/rooms/third-party-join?tabs=MTRW)  
If you are not seeing the join button for 3rd party links, make sure these are set as follows  
`DeleteComments: $false`  
`ProcessExternalMeetingMessages: $true`  


## URL Rewriting Tools  
See Step 3A here: [Third-party meetings on Teams Rooms](https://learn.microsoft.com/en-us/microsoftteams/rooms/third-party-join?utm_source=chatgpt.com&tabs=MTRW#step-3a-configure-url-rewrite-policies-to-not-modify-third-party-meeting-links).  
And format for Safe links Exclusions here: [Safe Links Formatting](https://learn.microsoft.com/en-us/defender-office-365/safe-links-about#entry-syntax-for-the-do-not-rewrite-the-following-urls-list)   
  
If you have a tool (e.g. Microsoft Defender) that rewrites URLs in message bodies, you will have to exclude zoom.us (e.g.) or else the device will not recognize join links.  
  
*Microsoft Defender > Email & Collaboration > Policies & rules > Threat policies Safe Links > Protection Settings > Edit > Do not rewrite the following URLs in email*  
e.g. `*.zoom.us/*` covers all the interations you would need for zoom links. 
 

## Allowing 3rd Party Joining (Setting on Device itself)  
If you are not seeing the join button, see Step 4 here: [Third-party meetings on Teams Rooms](https://learn.microsoft.com/en-us/microsoftteams/rooms/third-party-join?utm_source=chatgpt.com&tabs=MTRW#step-4-enable-third-party-meetings-on-your-teams-rooms-devices)  
This can only be done on the device itself (unless you have *Teams Room Pro* licensing).  
*Microsoft Teams Room device console > More > Settings >*   
[enter the device administrator username and password]  
*Meetings tab > Cisco Webex, Zoom, etc*  

## Password Change

Rooms should be exempt from regular password changes. Microsoft by default does not change passwords at the org level. If your org has password expiration configured, change the room user's password policy setting to not expire.\
See Microsoft info here: [Set password to never expire](https://learn.microsoft.com/en-us/microsoft-365/admin/add-users/set-password-to-never-expire?view=o365-worldwide)\
See IT Automator useful script here: [User Password Expiration Manager](https://github.com/ITAutomator/UserPasswordExpiration)

## Teams Rooms Visibility

Teams Rooms (by themselves) show limited information (just the subject) on their scheduling panels.

*Subject*\
The room panels show the subject. For sensitive topics this may be undesirable. The above settings can replace the subject with the name of
the organizer.

*Message Body and Attachments*\
Message body and attachments *do not* appear on the scheduling or control panels, either before or during the meeting.\
For this reason, hiding the body and attachments only becomes necessary if the room calendar is visible to employees outside the room who should not normally see the
information.\
*Warning: stripping the body will cause 3rd party join links to fail.*

## Calendar Visibility via Sharing and Permissions

Notwithstanding the `Set-CalendarProcessing` settings described above, you can also adjust the room calendar's visibility via the **Sharing and
Permissions** settings that are common to all user calendars.  

*`Set-CalendarProcessing` (covered above)*\
This strips and permanently changes the calendar items. Only used in rooms. This is independent of *sharing and permissions*.  

*Sharing and Permissions (covered here)*\
**Outlook > Calendar > Properties > Sharing and Permission**\
This controls how others can see the calendar from their Outlook. All
users have this control -- by default it shares free-busy information.\
Microsoft information is here: [Sharing Calendars](https://support.microsoft.com/en-us/office/share-an-outlook-calendar-as-view-only-with-others-353ed2c1-3ec5-449d-8c73-6931a0adab88)

The 'people in my organization' is set to (as a default) *free-busy* (aka *can view when I'm busy*) to the entire company. This is the minimum setting.

You can share the calendar with specific people by adding user entries
and using the '*Can view titles and locations*' setting. This will share
subjects but not message body or attachment contents.

# Powershell Command (Sample)
  
If you don't want to use the script, here are some helpful commands you can try.  

```powershell
# Required Module: uninstall
Uninstall-Module ExchangeOnlineManagement -AllVersions
 
# Required Module: installs, loads, 
Install-Module ExchangeOnlineManagement  # installs from online
Import-Module ExchangeOnlineManagement   # loads an installed module into memory
 
# Required Module: Check for new version
Get-Module ExchangeOnlineManagement -ListAvailable # shows whats locally installed
Find-Module ExchangeOnlineManagement               # shows whats online
Update-Module ExchangeOnlineManagement             # updates (if installed)
 
# This connects to 365 using module ExchangeOnlineManagement
Connect-ExchangeOnline         # Connects to tenant
Connect-ExchangeOnline –Device # (PS7 only) Use open web page with pasted code.
Disconnect-ExchangeOnline    # Disconnect from tenant
 
# Get a list of room mailboxes in the tenant
Get-Mailbox -RecipientTypeDetails RoomMailbox |
    Select-Object DisplayName, PrimarySmtpAddress

# Show all rooms’ settings
$mbps = $mbs.PrimarySmtpAddress | Get-CalendarProcessing
$mbps | Select-Object Identity, AutomateProcessing, AddOrganizerToSubject, AllowRecurringMeetings, DeleteAttachments, DeleteComments, DeleteSubject, ProcessExternalMeetingMessages  | Out-GridView
# Done

# Gets a room's settings
Get-CalendarProcessing -Identity 'room@domain.com' | Select-Object DisplayName, AutomateProcessing, AddOrganizerToSubject, AllowRecurringMeetings, DeleteAttachments, DeleteComments, DeleteSubject, ProcessExternalMeetingMessages
# Done

# Sets a room to Microsoft's default recommended settings
Set-CalendarProcessing -Identity 'RoomMailbox' ` 
  -AutomateProcessing AutoAccept ` 
  -AddOrganizerToSubject $false `
  -AllowRecurringMeetings $true `
  -DeleteAttachments $true `
  -DeleteComments $false ` 
  -DeleteSubject $false ` 
  -ProcessExternalMeetingMessages  $true

```

