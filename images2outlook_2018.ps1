## DIRECTORIES
##
$StudentDirectory = '\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Students\96x96-Outlook SMALL SIZE'
$StaffDirectory = '\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Staff\96x96Outlook-10kb\All Staff'
#$StaffDirectory = '\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Staff\96x96Outlook-10kb'
$body = @()

## FUNCTIONS
import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto
Function Get-DistinguishedNameContact ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
  
Function Get-DistinguishedNameUser ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}  


#Staff
Dir $StaffDirectory\*.jpg | export-csv $StaffDirectory\photo-list.csv 

import-csv $StaffDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$DistName = Get-DistinguishedNameUser($Email)
$FileName = $_.Name
Import-RecipientDataProperty -Identity $DistName -Picture -FileData ([Byte[]]$(Get-Content -Path $StaffDirectory\$FileName -Encoding Byte -ReadCount 0))
Write-Host "Photo set for AD User Account:"  $Email
$body+= "Photo Updated for $Email"}

#Students Mail Contact
Dir $StudentDirectory\*.jpg | export-csv $StudentDirectory\photo-list.csv 

import-csv $StudentDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$Alias = $Email.Substring(0,6)
#$DistName = Get-DistinguishedNameUser($Email)
$FileName = $_.Name
Import-RecipientDataProperty -Identity $Alias -Picture -FileData ([Byte[]]$(Get-Content -Path $StudentDirectory\$FileName -Encoding Byte -ReadCount 0))
Write-Host "Photo set for AD Contact:"  $Alias
$body+= " Mail Contact Photo Updated $Alias"}


#Students Login Accounts
import-csv $StudentDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$samAccountName = $Email.Substring(0,6)
$userPrincipalName = $samAccountName+'@bishanoi.com'
#$DistName = Get-DistinguishedNameContact($Email)
$FileName = $_.Name
$photo = [byte[]](Get-Content $StudentDirectory\$FileName -Encoding byte)
Set-ADUser -Identity $samAccountName -Replace @{thumbnailPhoto=$photo}
Write-Host "Photo set for Student Login Account:"  $samAccountName
$body+= "Student Login Account Photo Updated $samAccountName"}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "IT Info <it-info@bishanoi.com>" -Subject "AD & Outlook Update Photos 2016-2017" -Body $body -Smtpserver "mail.bishanoi.com"


## DIRECTORIES
##
$StudentDirectory = '\\nas\M\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Students\96x96Outlook-10kb'
$StaffDirectory = '\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Staff\96x96Outlook-10kb'
$body = @()

## FUNCTIONS
import-module ActiveDirectory
Function Get-DistinguishedNameContact ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
  
Function Get-DistinguishedNameUser ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}  


#Staff
Dir $StaffDirectory\*.jpg | export-csv $StaffDirectory\photo-list.csv 

import-csv $StaffDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$DistName = Get-DistinguishedNameUser($Email)
$FileName = $_.Name
Import-RecipientDataProperty -Identity $DistName -Picture -FileData ([Byte[]]$(Get-Content -Path $StaffDirectory\$FileName -Encoding Byte -ReadCount 0))
Write-Host "Photo set for AD User Account:"  $Email
$body+= "Photo Updated for $Email"}

#Students Mail Contact
Dir $StudentDirectory\*.jpg | export-csv $StudentDirectory\photo-list.csv 

import-csv $StudentDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$Alias = $Email.Substring(0,6)
#$DistName = Get-DistinguishedNameUser($Email)
$FileName = $_.Name
Import-RecipientDataProperty -Identity $Alias -Picture -FileData ([Byte[]]$(Get-Content -Path $StudentDirectory\$FileName -Encoding Byte -ReadCount 0))
Write-Host "Photo set for AD Contact:"  $Alias
$body+= " Mail Contact Photo Updated $Alias"}


#Students Login Accounts
import-csv $StudentDirectory\photo-list.csv  |
ForEach-Object {
$Email = $_.BaseName
$samAccountName = $Email.Substring(0,6)
$userPrincipalName = $samAccountName+'@bishanoi.com'
#$DistName = Get-DistinguishedNameContact($Email)
$FileName = $_.Name
$photo = [byte[]](Get-Content $StudentDirectory\$FileName -Encoding byte)
Write-Host "Photo set for Student Login Account:"  $samAccountName
$body+= "Student Login Account Photo Updated $samAccountName"}

$body = $body | out-string

Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "Heidi Hendry <it-info@bishanoi.com>" -Subject "AD & Outlook Update Photos 2017-2018" -Body $body -Smtpserver "mail.bishanoi.com"
