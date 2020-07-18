import-module ActiveDirectory

Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}  

# Create the AD User and depending on "Year" create it in the relevant OU Primary or Secondary
$csv = read-host("What is the full path location of the CSV file?")
import-csv $csv | ForEach-Object {
$UserID = $_.cn
$Email = ($UserID + "@company.net")
$DN = Get-DistinguishedNameUser $UserID
$ID = $_.cn

Set-ADUser -Identity $DN -EmailAddress $Email -UserPrincipalName $Email



}
