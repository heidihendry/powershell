import-module ActiveDirectory
Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 

import-csv fixhomedirectory.csv | ForEach-Object {
$UserID = $_.samAccountName
$HomeDir = ('\\192.168.150.55\studentpro$\' + $UserID)
$DN = Get-DistinguishedNameUser $UserID
Set-ADUser -Identity $DN  -HomeDirectory $HomeDir -HomeDrive ('S:')
}