import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

# Get the user's Distinguished Name 
Function Get-DistinguishedNameContact ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
  
Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}   

#Update fields on the AD User
import-csv “\\mail\c$\scripts\addNewEmail\addNewEmail.csv” | ForEach-Object {

$NewEmail = $_.new_email
$DN = Get-DistinguishedNameContact $_.current_email

Set-ADUser $DN -Add @{ProxyAddresses="smtp:$NewEmail"}
Write-Host "AD User New Email Address added:"   $_.current_email $_.newemail
}



