# Microsoft Exchange Toolbox V1

A menu-based toolbox built in Powershell to efficiently perform Admin Tasks for Mailboxes and Calendars for Service Desk Analysts and Deskside Engineers in Microsoft Exchange

<a href='https://ko-fi.com/Z8Z6E0CY0' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi2.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>

![Exchange Toolbox V1 Home Page 01](https://user-images.githubusercontent.com/8764255/180898335-616002c2-3bbd-4ea1-bcdf-556b82e91570.png)


# Prerequisites

To run this script you will need to be able to run unsigned Powershell scripts on your machine. If you are unable to download or run these on your device then you may need to check with your corporate Infrastructure team to enable this on your device. This script was tested using a Remote Packaged Application version of Windows Powershell as I was unable to get this working locally on my Dell Windows Laptop. I was unable to change this so stored this file on a network drive which was on the same server as the .PS1 script file

- A Microsoft Admin Account for your Organisational Microsoft Exchange Server
- Windows Powershell (With No Restriction on Code Execution)
- Permission to run this on your network

# How to Launch the Application (First Time Use)

This guide will walk you through the usage with screenshots of the Exchange Toolbox and how to launch the script when using for the first time. This guide can also be used as reference if you need later. 

There is a separate guide for an Example Scenario of “Adding a User to A Shared Calendar”. That guide breaks down that process from start to finish with a way to check your work as well with “List all Users that have Delegate Access to a Mailbox”. 

The exact link to the script due to an issue with it running from the S Drive at present but the process to launch this application will remain the same.

1. Download and install .ps1 script to a location such as C:\MicrosoftExchangeToolbox

2. Open up Windows Powershell

3. Once Windows Powershell is open, Change Directory to get to the Network location of the script. 

4. Then you need to Launch the application from Powershell which you can do by using the below command:

<code>./ExchangeToolboxV1.ps1</code>

Please Note: This command asks Powershell to run the script in the directory you are currently working in.

5. You have now launched the application and will be asked to login using your full ADM email address for your organisation e.g. adm-joeb@work.com and current password before being presented with the Main Menu.

![Exchange Toolbox V1 Home Page 02](https://user-images.githubusercontent.com/8764255/180899049-1c8a0e17-1ad5-423b-b3d2-656e4beb2ff6.png)

# FAQ

### What doesn't this do?

This toolbox doesn't have an Active Directory functions in it so you won't be able to add or remove people from Distribution Lists etc or to Security Groups.

### Can this break something?

The Powershell commands used in this toolbox are safe to use and not run a script unless all required information has been entered. There are no commands in here to delete or create any Mailboxes or Calendars, only adding, removing or listing them. With that said I take no responsibility for misuse of this tool but as you have to use your Admin Credentials, any changes to mailboxes will be logged by your organisation infrastructure team so make sure you've asked for permission to use this in a production environment.

### Can I change or edit the menu options easily?

You can change any of the Menu Texts in the code to something more applicable to your organisation but please try and link back here incase of any updates or refinements to this tool in future.

### How do I donate to you?

You can buy me a coffee on https://ko-fi.com/reprodev and check out my other projects too. Come say hi and let me know you came from GitHub

<a href='https://ko-fi.com/Z8Z6E0CY0' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi2.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>

# Background

In my role as a Service Desk Analyst mixing a bit of First Line and Second Line at an organisation, one of the tasks was to add and remove Executive Assistants to Mailboxes and Calendars. This was usually a pretty straightforward task but would either be done in Microsoft Exchange Admin click by click or for Calendars in Powershell line by line. We usually got these requests once every week or so and as a team we were able to action these requests within our SLA and normal workload. Howeever, we were moving to a new solution as a project throughout the organisation and from a VDI environment with Thin Client Laptops to one where each user now had a Fat Client Dell Laptop instead.

During the project and after the move, some users found that they may have lost access to the mailbox or calendars in the way they were set up within the old VDI environment. We started to get a lot of tickets for people reporting they had lost access to these calendars or mailboxes even after their Manager had sent a Sharing Request again.

This caused a lot of manual work for the Service Desk team as a whole and so to try and automate this process I looked into creating a basic set of Powershell commands in an easier to use package. I had created something similar for myself and will share that soon but that was something which didn't do anything Mailbox or Calendar related.

Looking into it some users had lost permission to calendars but others still had this access. If they re-added the calendars in Outlook manually they would find them there but as some were still getting to grips with the new devices we would need to try another approach. The idea was to remove and then re-add them in an efficient way and that was what the first version of this tool did.

After liaising with our Infrastructure team and getting permission to use this on the network in production I created a version to test with the other analysts and wider team. Through user testing and feedback it became the version you see now which still needs functions added to it but fulfilled the need for a quick way to work on Calendars for our team.

## References:

### How to Make a Menu in Powershell

https://www.youtube.com/watch?v=avxwjext5f4

### Mailbox Powershell Functions

https://lazyadmin.nl/powershell/get-mailbox-permissions-with-powershell/

http://woshub.com/list-mailboxes-user-has-access-to/

https://o365reports.com/2021/05/25/most-useful-powershell-cmdlets-to-manage-exchange-online-mailboxes/

https://theitbros.com/get-mailbox/

https://docs.microsoft.com/en-us/powershell/module/exchange/?view=exchange-ps#mailboxes

# Finally!

Please go ahead and follow me for more and feel free to comment, fork, clone and use as you wish. If you've found a way to make this process better or if there is anything in here that needs to be amended to make it flow better. I'd love to hear from anyone that has given this a try.

<a href='https://ko-fi.com/Z8Z6E0CY0' target='_blank'><img height='36' style='border:0px;height:36px;' src='https://cdn.ko-fi.com/cdn/kofi2.png?v=3' border='0' alt='Buy Me a Coffee at ko-fi.com' /></a>

Yours technically,
ReproDev
