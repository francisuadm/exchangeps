> > > # Exchange Powershell Script

#### Displaying current mailbox calendar folder permission.
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

