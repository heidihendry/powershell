import-module ActiveDirectory

Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}  


import-csv “c:\scripts\ExpiredPasswords\changepwdexpiryrule.csv” | ForEach-Object {
$DN = $_.DistinguishedName
Set-ADUser $DN -PasswordNeverExpires $false
#Set-ADUser $DN -ChangePasswordAtLogon $true
Write-Host $DN
}