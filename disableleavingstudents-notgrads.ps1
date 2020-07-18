import-module ActiveDirectory
$body = @()
# Get the user's Distinguished Name 
Function Get-DistinguishedNameContact ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
  
# $DN = Get-DistinguishedNameContact $_.mail  
  
Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}   



import-csv “\\mail\script$\SIMSImport\LeavingStudents\leavers-notgrads.csv” | ForEach-Object {
$DN = Get-DistinguishedNameUser $_.cn 
Disable-ADAccount -Identity $DN
Move-ADObject -Identity $DN "OU=Leavers - Student Accounts,OU=Leaving Students,OU=Leavers,OU=AllUsers,OU=BIS-HN,DC=company,DC=com"
Write-Host "User Account Disabled:"  $_.cn $_.givenName $_.sn
$body+= "User Account Disabled: $DN"
$DNContact = Get-DistinguishedNameContact $_.Email
Disable-MailContact -Identity $DNContact -Confirm:$false
Move-ADObject -Identity $DNContact "OU=Leaving Student Contacts,OU=Leaving Students,OU=Leavers,OU=AllUsers,OU=BIS-HN,DC=company,DC=com"
Write-Host "Mail Contact Disabled:"  $_.cn $_.givenName $_.sn
$body+= "Mail Contact Disabled: $DNContact"
}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@company.com>" -To "Heidi Hendry <it-info@company.com>" -Subject "Disable Leaving Non-Grad Students" -Body $body -Smtpserver "mail.company.com"

