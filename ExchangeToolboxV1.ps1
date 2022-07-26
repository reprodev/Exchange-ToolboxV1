######################################################################################
# Exchange Toolbox V1.0 - Released 26/07/2022                                        #
#                                                                                    #
# Script Created by ReproDev:   https://github.com/reprodev/Exchange-ToolboxV1/      #
# Released Under MIT Licence                                                         # 
# Check out other projects :    https://github.com/reprodev/                         #
# Why not buy me a coffee? :    https://ko-fi.com/reprodev                           #
#                                                                                    #
######################################################################################

# Credential Function Loaded before Global Variable

function Update-Creds
{
# This code block will remove the User Credential saved for this session if you've logged in with the wrong credentials
Clear-Host
Write-Host "Re-enter your credentials..."
$global:UserCredential = Get-Credential
#[console]::beep(800,1000)
Write-Host "You're logged back in. Returning to the Previous Menu..."
Pause
}

# ------------------------ LIST OF GLOBAL VARIABLES  -----------------------------

# Get-Credential cmdlet and asks for the user to log in with their credentials and then saves them as a secure.securitystring new object 
# until the app is closed

$global:UserCredential = Get-Credential

# --------------------------- LIST OF FUNCTIONS  -------------------------------

# ------- FUNCTIONS FOR MAILBOX ADMIN --------------

# This function adds a single user to a Mailbox and gives Full Access as the Default permission
function Add-Box-Single
{Clear-Host
  [string]$UserNameMailAdd = Read-Host "Enter the display name, email address or AD Username of the USER that needs access to a Mailbox. By Default, they will all be given Full Access to the Mailbox"
  [string]$MailBoxAdd = Read-Host "Enter the display name, email address or AD Username of the MAILBOX they need access to"
  #$PermissionMailAdd = Read-Host "What kind of permissions does the user need e.g. FullAccess, SendAs"
  Show-Session-Global
  Add-MailboxPermission -Identity $MailBoxAdd -User $UserNameMailAdd -AccessRights "FullAccess" -InheritanceType All
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
}

# This function removes fully removes a single user from a mailbox
function Remove-Box-Single
{Clear-Host
  [string]$UserNameMailRem = Read-Host "Enter the display name, email address or AD Username of the USER that you want to remove from a Mailbox"
  [string]$MailBoxRem = Read-Host "Enter the display name, email address or AD Username of the MAILBOX"
  Show-Session-Global
  Remove-MailboxPermission -Identity $MailBoxRem -User $UserNameMailRem -AccessRights "FullAccess" -InheritanceType All
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
}
# This function lists all users who have access to a mailbox
function ShareBoxList
{Clear-Host
  [string]$MailboxShCheck = Read-Host "Enter the display name, email address or AD Username of the MAILBOX to list the users who have access and their permission"
  Show-Session-Global
  Get-MailboxPermission -Identity $MailboxShCheck | Select-Object User,Identity,AccessRights | Format-List
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
}
# This function lists all the mailboxes a single user has delegate access to and the type of permission formatted in a list
function UserBoxList
{Clear-Host
  Write-Host "Please Note: THIS FUNCTION WILL ONLY WORK WITH THEIR EMAIL ADDRESS. Once you have entered the details. This searches all of Exchange and could take up to an hour to complete. A sound will play when the script has finished so please wait."
  [string]$MailboxUserCheck = Read-Host "Enter the EMAIL ADDRESS of the USER to list all the mailboxes they have delegate access to"
  Show-Session-Global
  # Get-Mailbox | Get-MailboxPermission -Identity $MailboxUserCheck | Format-List - Can be used to format this as a list instead of below but will cut off long Display Names and take just as long
  # Please note this one will take a long time and may look like it's stalled and show errors of duplicate email address the below [console]beep sound will play when the script has finished so you can do something else while you wait
  #Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -User $MailboxUserCheck | Format-List
  Get-Mailbox -RecipientTypeDetails UserMailbox,SharedMailbox -ResultSize Unlimited | Get-MailboxPermission -User $MailboxUserCheck | Select-Object User,Identity,AccessRights | Format-List
  #$global:Sounds
  [console]::beep(800,1000)
  Pause
  Get-PSSession | Remove-PSSession
}
# This function adds multiple mailboxes from a single user in a loop until the user types "N". 
# Type in the User once and then the you will be prompted for each Mailbox without having to retype the user
function Add-Box-Multi
{Clear-Host
  Show-Session-Global
  [string]$UserNameAddMailMulti = Read-Host "Enter the display name, email address or AD Username of the USER you want to add to multiple Mailboxes e.g. Staff Mover"
  [string]$UserNameAddMailLoop = 'Do you want to add multiple Mailboxes for this user. By default, they will all be given Full Access to the Mailbox [Y/N]'
  do 
  {
      $response = Read-Host -Prompt $UserNameAddMailLoop
      if ($response -eq 'y'){[String]$UserNameAddMailMultiBox = Read-Host "Enter the display name, email address or AD Username of the MAILBOX you want to add this user to, e.g. New Team Shared Inbox"
      {
      }
      Add-MailboxPermission -Identity $UserNameAddMailMultiBox -User $UserNameAddMailMulti -AccessRights "FullAccess" -InheritanceType All #The FullAccess can be changed to ReadPermission but you'll usually be adding just as FullAccess which is why this is the default
      #$global:Sounds
    } 
  } until ($response -eq 'n')
  Pause
  Get-PSSession | Remove-PSSession
}
# This function removes multiple mailboxes from a single user in a loop until the user types "N". 
# Type in the User once and then the you will be prompted for each Mailbox without having to retype the user
function Remove-Box-Multi
{Clear-Host
                    
  Show-Session-Global
  [string]$UserNameRemMailMulti = Read-Host "Enter the display name, email address or AD Username of the USER you want to remove from Mailboxes"
  [string]$UserNameRemMailLoop = 'Do you want to remove multiple Mailboxes from this user [Y/N]'
 do 
 {
     $response = Read-Host -Prompt $UserNameRemMailLoop
     if ($response -eq 'y'){$UserNameRemMailMultiBox = Read-Host "Enter the display name, email address or AD Username of the MAILBOX you want to remove from this user, e.g. Previous Team Shared Mailboxes"
     {
     }
     Remove-MailboxPermission -Identity $UserNameRemMailMultiBox -User $UserNameRemMailMulti -AccessRights "FullAccess" -InheritanceType All
     #$global:Sounds
   } 
 } until ($response -eq 'n')
 Pause
 Get-PSSession | Remove-PSSession
}

# ------------ FUNCTIONS FOR CALENDAR ADMINISTRATION --------------

# This function lists all the users that have delegate access to a calendar and the type of permission formatted in a list
function UserCalList
  {Clear-Host
  [string]$CalendarCheck = Read-Host "Enter the display name, email address or AD Username of the CALENDAR to check delegate permissions for"
  Show-Session-Global
  Get-MailboxFolderPermission -Identity $CalendarCheck":\calendar"
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
}
# This function lets you add a single user to a calendar and the type of permissions they need
function Add-Cal-Single
  {Clear-Host
  [string]$CalendarAdd = Read-Host "Enter the display name, email address or AD Username of the CALENDAR"
  [string]$UserNameAdd = Read-Host "Enter the display name, email address or AD Username of the USER that needs access to this calendar"
  [string]$PermissionAdd = Read-Host "What kind of permissions does the user need e.g. Owner, Editor, PublishingEditor"
  Show-Session-Global
  Add-MailboxFolderPermission -Identity $CalendarAdd":\calendar" -User $UserNameAdd -AccessRights $PermissionAdd
  #$global:Sounds
  Pause            
  Get-PSSession | Remove-PSSession
}
# This function lets you remove a single user to a calendar and the type of permissions they need
function Remove-Cal-Single
  {Clear-Host
  [string]$CalendarRem = Read-Host "Enter the display name, email address or AD Username of the CALENDAR"
  [string]$UserNameRemove = Read-Host "Enter the display name, email address or AD Username of the USER to remove from this calendar"
  Show-Session-Global
  Remove-MailboxFolderPermission -Identity $CalendarRem":\calendar" -User $UserNameRemove
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
}
# This function lets you change a single user's access to a calendar and the type of permissions they have
function Set-Cal-Single
  {Clear-Host
  [string]$CalendarSet = Read-Host "Enter the display name, email address or AD Username of the CALENDAR"
  [string]$UserNameSet = Read-Host "Enter the display name, email address or AD Username of the USER changing access to this calendar"
  [string]$PermissionSet = Read-Host "What kind of permissions does the USER need now e.g. Owner, Editor, PublishingEditor"
  Show-Session-Global
  Set-MailboxFolderPermission -Identity $CalendarSet":\calendar" -User $UserNameSet -AccessRights $PermissionSet
  #$global:Sounds
  Pause
  Get-PSSession | Remove-PSSession
} 
# This function removes multiple calendars from a single user in a loop until the user types "N". 
# Type in the User once and then the you will be prompted for each Calendar without having to retype the user
function Remove-Cal-Multi
  {Clear-Host
  Show-Session-Global
  [string]$UserNameRemMulti = Read-Host "Enter the display name, email address or AD Username of the USER to remove from calendars e.g. Staff Mover"
  [string]$UserNameRemLoop = 'Do you want to remove multiple calendars from this user [Y/N]'
  do 
  {
      $response = Read-Host -Prompt $UserNameRemLoop
      if ($response -eq 'y'){[String]$UserNameRemMultiCals = Read-Host "Enter the display name, email address or AD Username of the calendar you want to remove e.g. Team Calendar"
      {
      }
      Remove-MailboxFolderPermission -Identity $UserNameRemMultiCals":\calendar" -User $UserNameRemMulti
      #$global:Sounds
    } 
  } until ($response -eq 'n')
  Pause
  Remove-PSSession $Session}


# ---------- Functions for Session Management --------------

# This function starts a session in Exchange Online to carry out the function for carrying out Powershell actions
function Show-Session-Global
{ Clear-Host
  # Writes a message on screen to confirm that the login script to create a new session has started
  Write-Host "Running the requested script..." -ForegroundColor DarkBlue -BackgroundColor Gray
  # This sets the variable Session in the global scope and is the command that you would run before commands usually when you do something with Exchange in Powershell
  $global:Session =
    New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $global:UserCredential -Authentication "Basic" -AllowRedirection;
    # The below line then imports that session using the Variable we have just created to allow us to send the output of this function into the option we need
    # This is why the Menu Options contain these functions so this login part can be changed just here if this changes in future  
    Import-PSSession -Session $global:Session
}
# ---------- Functions for Menus and Submenus ----------
function Show-MainMenu-Toolbox
{
   # This is the create function to display the one level flat Menu and the Opt variables are called later for the the Write-Hosts.
   # Just change the text in the Double Speech marks and it will change in the menu.
   # You can add as many options as you want here make sure you wrap the new string in 2 round brackets above and below.
   # Leave Title as it is in the String Variable here and also below as you will call this later in the Do for User Input later on as -Title.
   # Each of the code blocks have variables that should relate to the function of that section for easier readibility and to make any 
   # changes to this script easier in the future.
    param (
        [string]$MainTitle = "My Menu"
          )
          (
        [string]$MainOpt1 = "Exchange Mailbox Administration"
          )
          (
        [string]$MainOpt2 = "Exchange Calendar Administration"
          )
          (
        [string]$MainOpt3 = "Show All Admin Programs"
          )
          (
        [string]$MainOpt4 = "Update Your Saved Credentials"
          )
    Clear-Host
    Write-Host "================== $MainTitle ==================" -ForegroundColor DarkBlue -BackgroundColor Gray
    Write-Host ""
    Write-Host " 1: $MainOpt1"
    Write-Host " 2: $MainOpt2"
    Write-Host " 3: $MainOpt3"
    Write-Host " 4: $MainOpt4"
    #Write-Host " 5: $MainOpt5"
    #Write-Host " S: Turn On/Off Sounds"
    Write-Host ""
    Write-Host "=======Type 'Q' to Quit The Application========"
}
function Show-SubMenu-Calendar
{
   # This is the create function to display the calendar Sub Menu and the CalOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to CalTitle so that this menu has it's own Title
    param (
        [string]$CalTitle = "My Menu"
          )
          (
        [string]$CalOpt1 = "List all Users that have Delegate Permissions on a Calendar"
          )
          (
        [string]$CalOpt2 = "Add a User to a Calendar"
          )
          (
        [string]$CalOpt3 = "Remove a User from a Calendar"
          )
          (
        [string]$CalOpt4 = "Change a User's Access Permissions to a Calendar"
          )
          (
        [string]$CalOpt5 = "Remove an Individual User From Multiple Calendars"
          )
          (
        [string]$CalOpt6 = "Update Your Saved Credentials"
          )
    Clear-Host
    Write-Host "================== $CalTitle ==================" -ForegroundColor DarkBlue -BackgroundColor Gray
    Write-Host ""
    Write-Host " 1: $CalOpt1"
    Write-Host " 2: $CalOpt2"
    Write-Host " 3: $CalOpt3"
    Write-Host " 4: $CalOpt4"
    Write-Host " 5: $CalOpt5"
    Write-Host " 6: $CalOpt6"
    #Write-Host " S: Turn On/Off Sounds"
    Write-Host ""
}
function Show-SubMenu-Mailbox
{
   # This is the create function to display the calendar Sub Menu and the MailOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to MailTitle so that this menu has it's own Title
    param (
        [string]$MailTitle = "My Menu"
          )
          (
        [string]$MailOpt1 = "Give a User Delegate Access to a Mailbox"
          )
          (
        [string]$MailOpt2 = "Remove a User's Delegate Access to a Mailbox"
          )
          (
        [string]$MailOpt3 = "List all Users that have Delegate access to a Mailbox"
          )
          (
        [string]$MailOpt4 = "List all Mailboxes a User has Delegate Access To"
          )
          (
        [string]$MailOpt5 = "Add an Individual User as Delegate to Multiple Mailboxes"
          )
          (
        [string]$MailOpt6 = "Remove an Individual User as a Delegate From Multiple Mailboxes"
          )
          (
        [string]$MailOpt7 = "Update Your Saved Credentials"
          )
    Clear-Host
    Write-Host "================== $MailTitle ==================" -ForegroundColor DarkBlue -BackgroundColor Gray
    Write-Host ""
    Write-Host " 1: $MailOpt1"
    Write-Host " 2: $MailOpt2"
    Write-Host " 3: $MailOpt3"
    Write-Host " 4: $MailOpt4"
    Write-Host " 5: $MailOpt5"
    Write-Host " 6: $MailOpt6"
    Write-Host " 7: $MailOpt7"
    #Write-Host " S: Turn On/Off Sounds"
    Write-Host ""
}
function Show-SubMenu-Full
{
   # This is the create function to display all the scripts in a Sub Menu and the FullOpt variables are called later for the the Write-Hosts.
   # The Title has also been changed in the Write-Host and Parameter to FullTitle so that this menu has it's own Title
    param (
        [string]$FullTitle = "My Menu"
          )
          (
        [string]$FullOpt1 = "Give a User Delegate Access to a Mailbox"
          )
          (
        [string]$FullOpt2 = "Remove a User's Delegate Access to a Mailbox"
          )
          (
        [string]$FullOpt3 = "List all Users that have Delegate access to a Mailbox"
          )
          (
        [string]$FullOpt4 = "List all Mailboxes a User has Delegate Access To"
          )
          (
        [string]$FullOpt5 = "Add an Individual User as Delegate to Multiple Mailboxes"
          )
          (
        [string]$FullOpt6 = "Remove an Individual User as a Delegate From Multiple Mailboxes"
          )
          (
        [string]$FullOpt7 = "List all Users that have Delegate Permissions on a Calendar"
          )
          (
        [string]$FullOpt8 = "Add a User to a Calendar"
          )
          (
        [string]$FullOpt9 = "Remove a User from a Calendar"
          )
          (
        [string]$FullOpt10 = "Change a User's Access Permissions to a Calendar"
          )
          (
       [string]$FullOpt11 = "Remove an Individual User From Multiple Calendars"
          )
          (
        [string]$FullOpt12 = "Update Your Saved Credentials"
          )
    Clear-Host
    Write-Host "================== $FullTitle ==================" -ForegroundColor DarkBlue -BackgroundColor Gray
    Write-Host ""
    Write-Host " 1: $FullOpt1"
    Write-Host " 2: $FullOpt2"
    Write-Host " 3: $FullOpt3"
    Write-Host " 4: $FullOpt4"
    Write-Host " 5: $FullOpt5"
    Write-Host " 6: $FullOpt6"
    Write-Host " 7: $FullOpt7"
    Write-Host " 8: $FullOpt8"
    Write-Host " 9: $FullOpt9"
    Write-Host "10: $FullOpt10"
    Write-Host "11: $FullOpt11"
    Write-Host "12: $FullOpt12"
    #Write-Host " S: Turn On/Off Sounds"
    Write-Host ""
}
# --------------------------------- END OF FUNCTION LIST -----------------------------------------------------
#
# Do Loop for Menu containing user switches and a nested looping system
# This also calls the Menus with a changed positional paremeter switch so you can change these next to e.g. "-MainTitle" and "-MailTitle"
# Use a new variable for each menu or it'll loop indefinitely or just have no menu e.g. $MainTile and $MailTitle
# You can add as many options as you want to each layer and then branch out from there
# Functions have been used so they can be called instead of code blocks in the User Switches for easier readibility and editing of the options later
do
{
    Clear-Host
    Show-MainMenu-Toolbox -MainTitle "Exchange Toolbox v1.0"
    $UserMain = Read-Host "Please enter your selection from the list"
    #$global:Sounds = [console]::beep(800,1000)
    switch ($UserMain)
     {
           "1" {
                
                do
                    {
                        Clear-Host
                        Show-SubMenu-Mailbox -MailTitle "Microsoft Exchange Mailbox Administration"
                        $UserSubMail = Read-Host -Prompt 'Enter 1 - 7 or B for Back'
                        Switch ($UserSubMail)
                         {
                            '1' {Add-Box-Single}
                            '2' {Remove-Box-Single}
                            '3' {ShareBoxList}
                            '4' {UserBoxList}
                            '5' {Add-Box-Multi}
                            '6' {Remove-Box-Multi}  
                            '7' {Update-Creds}
                         }
                         #pause
                    }
                    until ($UserSubMail -eq "b")
           } "2" {
               Clear-Host
               do
                {
                    Clear-Host
                    Show-SubMenu-Calendar -CalTitle "Microsoft Exchange Calendar Administration"
                    $UserSubCal = Read-Host -Prompt 'Enter 1 - 6 or B for Back'
                    Switch ($UserSubCal)
                     {
                        '1' {UserCalList}          
                        '2' {Add-Cal-Single}
                        '3' {Remove-Cal-Single} 
                        '4' {Set-Cal-Single}     
                        '5' {Remove-Cal-Multi}
                        '6' {Update-Creds}
                     }
                     #pause
                }
                until ($UserSubCal -eq "b")
           } "3" {
            Clear-Host
               do
               {
               Clear-Host
               Show-SubMenu-Full -FullTitle "Full List Of Applications"
               $UserSubFull = Read-Host -Prompt 'Enter 1 - 12 or B for Back'
               Switch ($UserSubFull)
                {
                   '1' {Add-Box-Single}        
                   '2' {Remove-Box-Single}
                   '3' {ShareBoxList}
                   '4' {UserBoxList}     
                   '5' {Add-Box-Multi}
                   '6' {Remove-Box-Multi}
                   '7' {UserCalList}
                   '8' {Add-Cal-Single}
                   '9' {UserBoxList}
                  '10' {Set-Cal-Single}
                  '11' {Remove-Cal-Multi}
                  '12' {Update-Creds}
                  }
                  #pause
                }
                until ($UserSubFull -eq "b")
          } "4" {Clear-Host
                 Update-Creds
          }"q" {
                return
           }
     }
     #pause
}
until ($UserMain -eq "q")
