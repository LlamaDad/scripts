#Connect to Exchange
https://github.com/LlamaDad/tools/AzureConnect.ps1

#Get-Major Version of Powershell
$PS = $PSVersionTable.PSVersion.major.ToString()

#If Major Version of Powershell is less than 4, Load ActiveDirectory Module
If($PS -lt 4){
    Import-Module -Name Active Directory
}

Else {
    Write-Host "Powershell version is up to date"
}

#Prompt for Mailbox Type
$mbtype = Read-Host "Please select the type of Mailbox

1 - Humad Mailbox

2 - HMHS Mailbox
"
#Environmental Variables
If ($mbtype -eq "1"){
    $domain = "humad.com"
    Set-ADServersettings -RecipientViewRoot $domain
    $sub = 6
}

ElseIf ($mbtype -eq "2") {
    $domain = "hmhschamp.humad.com"
    Set-ADServersettings -RecipientViewRoot $domain
    $sub = 10
}

Else{
    Write-Host "Please proide a valid input"
    Get-PSSession | Remove-PSSession
    exit
}

Write-Host "Getting objects with Full Access to Shared Mailboxes"
$groups = Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize:Unlimited | Get-MailboxPermissions | Where-Object {$_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false}

Write-Host "Groups with Full Access Recieved, now getting group membership"

$up = $env:USERPROFILE
$dt = $up + "\Desktop\"
$now = Get-Date
$path = "MBFA_" + $now.ToString('MMddyy') + $domain + ".csv"
$file = $dt + $path

ForEach($g in $groups)
    {
        $p = $g.User.Substring($sub)
        $q = $g.Identity
        Try{
            $m = Get-AdGroupMember -Identity $p
            $q, $p, $m.samaccountname, " " | Out-File -FilePath $file -Append
            }
        Catch{
            $q, $p, " " | Out-File -FilePath $file -append
            }
    }
Write-Host "Script Complete, please find output $file"
Get-PSSession | Remove-PSSession