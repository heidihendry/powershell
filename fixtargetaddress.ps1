import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

# Get the user's Distinguished Name 
Function Get-DistinguishedNameContact ($displayName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(displayName=$displayName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 

#Update targetAddress on the Student AD Contact
import-csv “\\mail.bishanoi.com\simsimport\AllStudents\allstudents.csv” | ForEach-Object {

$NewEmail = $_.cn + "@bishanoi.net"
$OldEmail = $_.cn + "@bishanoi.com"
$cn = $_.givenName + " " + $_.sn + " " + $_.cn
$DN = Get-DistinguishedNameContact $cn

Set-MailContact -Identity $DN  -targetAddress $NewEmail
Set-AdObject -Identity $DN -Remove @{targetAddress="SMTP:$OldEmail"} -Add @{targetAddress="SMTP:$NewEmail"}

Write-Host "AD Contact New Email Address replaced:"   $_.current_email $NewEmail
}