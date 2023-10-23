## Sharepoint Connection and Integration

Import-Module -Name PnP.PowerShell
gmo -l | ipmo

#Fill "" with Sharepoint url of table
Connect-PnPOnline -Url "" -Interactive

# Fill the "" with an List-code  in the pattern of this:  1111111a-0aa1-1aa1-a11a-111111111111  . Every number was changed to 1 and every letter disregarding if uppercase or lowercase has been changed to a lowercase a
$ListItems = Get-PnPListItem -List ""

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
     Write-Host " Approval Accepted. Specified Date Reached."
        # Perform the AD and Exchange operations
        # Change ** for your Active Directories Adress.
        Get-ADUser -Filter "UserPrincipalName -eq '$OldUPN'" -Server ** | Set-ADUser -UserPrincipalName $NewUPN -DisplayName $FirstName -sn $newLastName
        Write-Host " AD User has been changed for: "$OldUPN

        # Additional Exchange Online operations based on $OldUPN and $NewUPN
        # Change ** for your Exchange Onlines LDAP Adress
        Set-RemoteMailbox -Identity "UserPrincipalName -eq '$OldUPN'" -PrimarySmtpAddress $NewUPN -EmailAddressPolicyEnabled $false -DomainController **
        Write-Host " Exchange Online E-Mail has been changed for: "$OldUPN
    }
}

# Close the session after you've finished using it
Remove-PSSession -Session $Session
