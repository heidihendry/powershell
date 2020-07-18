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

#Update Graduate Students

#Update fields on the AD MailContact
import-csv “\\mail\script$\SIMSImport\LeavingStudents\leavers-aregrads.csv” | ForEach-Object {
$GradClass=$_.GradYear
$Office = ('Graduated ' + $GradClass)
$Dept = ('Graduated ' + $GradClass)
$DOA = $_.DOA
$Today = Get-Date
$DN = Get-DistinguishedNameContact $_.Email
Set-Contact -Identity $DN -company "Graduated from British International School, Hanoi" -department $Dept -Office $Office -title $Dept
Set-ADObject -Identity $DN -Description ('Graduated ' + $GradClass) -Replace @{info=("Date of Admission:" + $DOA + "  Updated on:" + $Today + "  Graduated " + $GradClass)} 
Write-Host "Grad Student MailContact fields updated:"  $DN
$body+="Grad Student MailContact fields updated:  $DN"
}


#Update fields on the AD User
import-csv “\\mail\script$\SIMSImport\LeavingStudents\leavers-aregrads.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$_.GradYear
$HomeDir = ('\\nas.COMPANY.com\studentpro$\' + $UserID)
$Office = ('Graduated ' + $GradClass)
$DN = Get-DistinguishedNameUser $_.cn
$Dept = ('Graduated ' + $GradClass)
Set-ADUser -Identity $DN -Description $Dept -Title $Dept -Department $Dept -AccountExpirationDate "Saturday, September 30, 2017 5:00 PM" -Office $Office -PasswordNeverExpires:$true -ChangePasswordAtLogon:$false -enabled:$true
Write-Host "Grad Student  User Account fields updated:"  $_.cn $_.givenName $_.sn
$body+="Grad Student  User Account fields updated: $UserID"
}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@COMPANY.com>" -To "Heidi Hendry <heidi.hendry@COMPANY.com>" -Subject "Update Grad Students" -Body $body -Smtpserver "mail.COMPANY.com"
