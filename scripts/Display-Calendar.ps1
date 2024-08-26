function Display-CalendarPermissions {
    # Remove existing PSSessions
    Get-PSSession | Remove-PSSession
    Clear-Host

    # Remove variables if they exist
    Remove-Variable -Name FullName -ErrorAction SilentlyContinue
    Remove-Variable -Name Domain -ErrorAction SilentlyContinue
    Remove-Variable -Name FolderCalendar -ErrorAction SilentlyContinue

    # Check if connected to Exchange Online
    if (-not (Get-PSSession | Where-Object { $_.ConfigurationName -eq 'Microsoft.Exchange' })) {
        # Check if ExchangeOnlineManagement module is installed
        if (-not (Get-InstalledModule -Name ExchangeOnlineManagement -ErrorAction SilentlyContinue)) {
            Write-Host "ExchangeOnlineManagement module is not installed. Installing..."
            Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
        } else {
            Write-Host "Updating ExchangeOnlineManagement module..."
            Update-Module -Name ExchangeOnlineManagement
        }

        # Connect to Exchange Online
        Write-Host "Connecting to Exchange Online..."
        Connect-ExchangeOnline
    } else {
        Write-Host "Already connected to Exchange Online."
    }

    # Prompt for identity of the mailbox to view
    Write-Host "[ Please enter first.last name of the identity mailbox to view: ] " -NoNewline
    $FullName = Read-Host
    $Domain = "@"+$EmailAddress.Split("@")[1]
    $FolderCalendar = $FullName+$Domain+":\Calendar"

    # Display the mailbox folder permission
    Write-Host
    Get-EXOMailboxFolderPermission -Identity $FolderCalendar | Select-Object -Property * | Out-GridView -Title "$FolderCalendar Properties Permission"

    # Disconnect from Exchange Online
    Disconnect-ExchangeOnline -Confirm:$false

    # Pause before exiting
    Read-Host -Prompt "Press Enter to exit"
}

# Call the function
Display-CalendarPermissions
