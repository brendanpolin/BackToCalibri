# Description
This Powershell Script will revert all of the default fonts in the O365 apps back to Calibri for all users . This includes Outlook new and old, Powerpoint, Excel, and Word.

# Requirements & Instructions
Create a folder called BackToCalibri on the C:\ drive
Put your Blank.potx, Normal.dotm, and NormalEmail.dotm in there. 

_The script will run without these, though powerpoint & Word won't configure though I'm not entirely sure about outlook_

Then simply run the script as administrator, and it will take care of the rest
```
.\BackToCalibri.ps1
```
_Sidenote: You can change the Blank.potx to any template if your org has a nice one that you would like to use_
## New Outlook Instructions
Log into your target organization's [Exchange Online Shell](https://learn.microsoft.com/en-us/powershell/exchange/connect-to-exchange-online-powershell?view=exchange-ps)

Run the following command to change the font for all users
```
Get-Mailbox -ResultSize unlimited | Set-MailboxMessageConfiguration -DefaultFontName "Calibri" -DefaultFontSize "10"
```

# Sources/Credits:

[https://www.pdq.com/blog/modifying-the-registry-users-powershell/](https://www.pdq.com/blog/modifying-the-registry-users-powershell/)

[https://github.com/j0eyv/ProactiveRemediations/blob/main/Remediate_OutlookFont_Calibri.ps1](https://github.com/j0eyv/ProactiveRemediations/blob/main/Remediate_OutlookFont_Calibri.ps1)

[https://www.shapechef.com/blog/change-the-default-font-in-powerpoint](https://www.shapechef.com/blog/change-the-default-font-in-powerpoint)
