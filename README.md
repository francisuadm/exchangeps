> # Exchange Powershell Script

> > ### Connecting to Exchange Online

```
Get-InstalledModule exchangeonlinemanagement | Format-List name,version, installedlocation
Update-Module -Name ExchangeOnlineManagement
Connect-Exchangeonline
```
#

> > > #### Displaying current mailbox calendar folder permission.
```
Get-PSSession | Remove-PSSession
Clear-Host
# First make sure variable are remove
Remove-Variable -Name FullName 
Remove-Variable -Name Domain 
Remove-Variable -Name FolderCalendar
# Prompt for identity of the mailbox to view
Write-Host "[ Please enter first.last name of the identity mailbox to view: ] " -NoNewline
$FullName = Read-Host
$Domain = "@"+$EmailAddress.Split("@")[1]
$FolderCalendar = $FullName+$Domain+":\Calendar"
# Display the mailbox folder permission
Write-Host
Get-EXOMailboxFolderPermission -Identity $FolderCalendar | Select-Object -Property * | Out-GridView -Title "$FolderCalendar Properties Permission"
```

> > > #### Add user as editor.
```
Get-PSSession | Remove-PSSession
Clear-Host
# First make sure variable are remove
Remove-Variable -Name FullName 
Remove-Variable -Name Domain 
Remove-Variable -Name FolderCalendar
Remove-Variable -Name editorEmail
# Prompt for identity of the mailbox to view
Write-Host "[ Please enter first.last name of the identity mailbox source: ] " -NoNewline
$FullName = Read-Host
# Code to add a user as an editor
Write-Host "Enter the user's email address to add as an editor:" -NoNewline
$editorEmail = Read-Host
$Domain = "@"+$EmailAddress.Split("@")[1]
$FolderCalendar = $FullName+$Domain+":\Calendar"
# Display the mailbox folder permission
Write-Host
Add-MailboxFolderPermission -Identity $FolderCalendar -User $editorEmail -AccessRights Editor -confirm:$false
```

> > > #### Add user as editor w/delegate.
```
Get-PSSession | Remove-PSSession
Clear-Host
# First make sure variable are remove
Remove-Variable -Name FullName 
Remove-Variable -Name Domain 
Remove-Variable -Name FolderCalendar
Remove-Variable -Name editorEmail
# Prompt for identity of the mailbox to view
Write-Host "[ Please enter first.last name of the identity mailbox source: ] " -NoNewline
$FullName = Read-Host
# Code to add a user as an editor with delegate
Write-Host "Enter the user's email address to add as an editor:" -NoNewline
$editorEmail = Read-Host
$Domain = "@"+$EmailAddress.Split("@")[1]
$FolderCalendar = $FullName+$Domain+":\Calendar"
# Display the mailbox folder permission
Write-Host
Add-MailboxFolderPermission -Identity $FolderCalendar -User $editorEmail -AccessRights Editor -SharingPermissionFlags Delegate -confirm:$false
```

> > > #### Remove user's calendar permission.
```
Get-PSSession | Remove-PSSession
Clear-Host
# First make sure variable are remove
Remove-Variable -Name FullName 
Remove-Variable -Name Domain 
Remove-Variable -Name FolderCalendar
Remove-Variable -Name editorEmail
# Prompt for identity of the mailbox to view
Write-Host "[ Please enter first.last name of the identity mailbox source: ] " -NoNewline
$FullName = Read-Host
# Code to add a user as an editor with delegate
Write-Host "Enter the user's email address to remove their calendar permission :" -NoNewline
$editorEmail = Read-Host
$Domain = "@"+$EmailAddress.Split("@")[1]
$FolderCalendar = $FullName+$Domain+":\Calendar"
# Display the mailbox folder permission
Write-Host
Remove-MailboxFolderPermission -Identity $FolderCalendar -User $editorEmail -confirm:$false
```
