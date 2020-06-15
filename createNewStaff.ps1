import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$body = @()
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




# Create the AD User 
import-csv “\\mail\scripts$\createNewStaff\newstaff.csv” | ForEach-Object {
$AccountName = $_.name	
New-ADUser $_.name -sAMAccountName $_.sAMAccountName –GivenName $_.givenName –Surname $_.sn -Path "OU=New Users,OU=Staff,OU=AllUsers,OU=BIS-HN,DC=bishanoi,DC=com"  -EmployeeID $_.employeeID  -description $_.description -title $_.title -department $_.department  -company "British International School, Hanoi"  
Write-Host "Staff AD account created:"  $_.cn $_.givenName $_.sn
$body+="Staff AD account created: $AccountName"
}




#Set Password
import-csv “\\mail\scripts$\createNewStaff\newstaff.csv” | ForEach-Object {
$PWD = ($_.employeeID + '!')
$AccountName = $_.name
$DN = Get-DistinguishedNameUser $_.cn
Set-ADAccountPassword -Identity $DN -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $PWD -Force) 

Write-Host "Account password set:"  $_.cn $_.givenName $_.sn
$body+="Account password set: $AccountName"
}


#Update fields on the AD User
import-csv “\\mail\scripts$\createNewStaff\newstaff.csv” | ForEach-Object {
$UserID = $_.cn
$HomeDir = ('\\nas.bishanoi.com\adminpro$\' + $UserID)
$DN = Get-DistinguishedNameUser $_.cn
$EmpID = $_.employeeID

Set-ADUser -Identity $DN -Description $_.description -Title $_.title -Department $_.department -HomeDirectory $HomeDir -HomeDrive ('S:') -EmailAddress $_.mail -Office $Office -Replace @{info=("Date of Start:" + $_.DOA + ". Updated on:" + $Today);userPrincipalName = $_.mail} -PasswordNeverExpires:$true -ChangePasswordAtLogon:$false -enabled:$true
Set-ADUser -Identity $DN -Replace @{employeeID=$EmpID}
Set-ADUser -Identity $DN -Add @{co="Vietnam"} -Company 'British International School, Hanoi'
#
#
#Update Groups
#
#
Add-ADGroupMember "CN=BIS_ADMIN,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=bishanoi,DC=com" –Member $DN #This line can be repeated for any additional groups if necessary
Write-Host "AD User Account fields updated:"   $_.givenName $_.sn $_.cn
$body+="AD User Account fields updated: $UserID" 
}

#Need to confirm that this works:
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ea SilentlyContinue
#Create the Mailbox
#Would be good if it could work out which Exchange Database ???
import-csv “\\mail\scripts$\createNewStaff\newstaff.csv” | ForEach-Object {

$DN = Get-DistinguishedNameUser $_.cn
$UserID = $_.cn
enable-Mailbox -Identity $DN  -Alias $_.cn -Database 'Teachers' -DisplayName ($_.name) -PrimarySmtpAddress $_.mail

Write-Host "Exchange Mailbox created:"  $_.cn $_.givenName $_.sn
$body+="Exchange Mailbox created: $UserID"
}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "IT Scripts <it-info@bishanoi.com>" -Subject "Create New Staff" -Body $body -Smtpserver "mail.bishanoi.com"

