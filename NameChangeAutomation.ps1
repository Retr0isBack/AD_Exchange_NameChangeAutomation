## Sharepoint Connection and Integration

Import-Module -Name PnP.PowerShell
gmo -l | ipmo

#Fill "" with Sharepoint url of table
Connect-PnPOnline -Url "" -Interactive

$ListItems = Get-PnPListItem -List "1122519b-0ea3-4be0-b25f-149292138820"

# Get the current date
$CurrentDate = Get-Date

# Build Connection with Exchange Online
#Replace ** with ConnectionUrl

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUrl **  -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

# Iterate through SharePoint list items
# The original list these parameters got imported from was written in german therefore the names.
foreach ($item in $ListItems) {
    $ItemDate = [datetime]::Parse($item["Umstellungsdatum"])
    $FirstName = $item["Vorname"]
    $oldLastName = $item["alterNachname"]
    $newLastName = $item["neuerNachname"]
    $domain = $item["Subdomain"]
    $OldUPN = $item["alteUPN"]
    $NewUPN = "$FirstName.$newLastName@$domain"
    $ItemStatus = $item["Status"]

    # Check if the date is in the past or equal to the current date and if the status is "Approved"
    if ($ItemDate -le $CurrentDate -and $ItemStatus -eq "Approved") {
        # Perform the AD and Exchange operations
        Get-ADUser -Filter "UserPrincipalName -eq '$OldUPN'" -Server ugfde555.de.ugfischer.com | Set-ADUser -UserPrincipalName $NewUPN -DisplayName $FirstName -sn $newLastName

        # Additional Exchange Online operations based on $OldUPN and $NewUPN
        Set-RemoteMailbox -Identity "UserPrincipalName -eq '$OldUPN'" -PrimarySmtpAddress $NewUPN -EmailAddressPolicyEnabled $false -DomainController ugfde445.de.ugfischer.com
    }
}

# Close the session after you've finished using it
Remove-PSSession -Session $Session
