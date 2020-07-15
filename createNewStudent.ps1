import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$PrimaryYear = "Foundation 1","Foundation 2","Foundation 3", "Year 01", "Year 02","Year 03","Year 04","Year 05","Year 06"
$EYCCLass = "Foundation 1","Foundation 2","Foundation 3"
$PrimaryYearClassGroups = "Year 01", "Year 02","Year 03","Year 04","Year 05","Year 06"
$SecondaryYear = "Year 07","Year 08","Year 09","Year 10","Year 11","Year 12","Year 13"
$SecondaryKS3 = "Year 07","Year 08","Year 09"
$SecondarySenior = "Year 10","Year 11","Year 12","Year 13"
$GradClassArr = @{"Year 13" = 2018;"Year 12"=2019;"Year 11"=2020;"Year 10"=2021;"Year 09"=2022;"Year 08"=2023;"Year 07"=2024;"Year 06"=2025;"Year 05"=2026;"Year 04"=2027;"Year 03"=2028;"Year 02"=2029;"Year 01"=2030;"Foundation 3"=2031;"Foundation 2"=2032;"Foundation 1"=2033}
$DOBArray = @{"January" = "01";"February" = "02";"March" = "03";"April"="04";"May"="05";"June"="06";"July"="07";"August"="08";"September"="09";"October"="10";"November"="11";"December"="12"}
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


# Create the AD User and depending on "Year" create it in the relevant OU Primary or Secondary
import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {

$GradClass=$GradClassArr.Get_Item($_.description)
if ($PrimaryYear -contains $_.description)
{
$ID = $_.cn	
New-ADUser -name $_.cn  –GivenName $_.givenName –Surname $_.sn -Path "OU=Primary Student Login Accounts,OU=Students,OU=AllUsers,OU=Company,DC=company,DC=com"  -DisplayName ($_.givenName + " " + $_.sn + " - " + $_.cn) -EmployeeID $_.employeeID -sAMAccountName $_.sAMAccountName -description $_.description -title $_.title -department $_.department  -company "British International School, Hanoi"  
$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=Student_Primary,OU=Security Groups,OU=Groups,OU=Company,DC=company,DC=com" –Member $DN
Write-Host "Primary account created:"  $_.cn $_.givenName $_.sn
$body+= "Primary account created: $ID"
}
elseif ($SecondaryYear -contains $_.description)
{
$ID = $_.cn
New-ADUser -name $_.cn  –GivenName $_.givenName –Surname $_.sn -Path "OU=Secondary Student Login Accounts,OU=Students,OU=AllUsers,OU=Company,DC=company,DC=com"  -DisplayName ($_.givenName + " " + $_.sn + " - " + $_.cn) -EmployeeID $_.employeeID -sAMAccountName $_.sAMAccountName -description $_.description -title $_.title -department $_.department  -company "British International School, Hanoi"
$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=Student_Secondary,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=company,DC=com" –Member $DN
Write-Host "Secondary account created:"  $_.cn $_.givenName $_.sn
$body+="Secondary account created: $ID"
}
else {
$ID = $_.cn
Write-Host "No Year Supplied, so no user created:" $_.cn $_.givenName $_.sn
$body+="No Year Supplied, so no user created: $ID"
}
}


#Create the Mail Contact
import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$Email = ($_.cn + "@company.net")
New-MailContact -name ($_.givenName + " " + $_.sn + " - " + $_.cn)  -alias $_.cn –FirstName $_.givenName –LastName $_.sn -ExternalEmailAddress $Email -PrimarySmtpAddress $Email -OrganizationalUnit "OU=Student Mail Contacts,OU=Students,OU=AllUsers,OU=BIS-HN,DC=company
,DC=com"  -Displayname ($_.givenName + " " + $_.sn + " - " + $_.cn)
Write-Host "Exchange Mail Contact created:"  $_.cn $_.givenName $_.sn
$body+="Exchange Mail Contact created:  $Email"
}

#Add the Mail Contact to the relevant Mail Distribution Group
# what if no class supplied? -> Need to have an Update Students for running 2 days after admission.

import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$Email = ($_.cn + "@company.net")
$GradClass=$GradClassArr.Get_Item($_.description)
$ID = $_.cn
$Class = $_.Reg
if ($PrimaryYearClassGroups -match $_.description)
{

Add-DistributionGroupMember -Identity ( "Class 0" + $_.Reg) -Member $Email
Write-Host "Primary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Primary Account $Email added to class $Class"
 }
elseif ($EYCClass -match $_.description)
{

Add-DistributionGroupMember -Identity ( "Class " + $_.Reg) -Member $Email
Write-Host "EYC Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="EYC Account $ID added to class $Class"
}

elseif ($SecondaryKS3 -match $_.description)
{

Add-DistributionGroupMember -Identity ( "Class 0" + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}

elseif ($SecondarySenior -match $_.description)
{

Add-DistributionGroupMember -Identity ( "Class " + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}
else {

Write-Host "No Class Supplied, so mail contact not added to Class group:" $_.cn $_.givenName $_.sn
$body+="No Class Supplied, so mail contact not added to Class group: $ID" }
}




#Set Password 
import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$DOBLong = $_.DOB
$DOBDay = $DOBLong.Substring(0,2) #first 2 characters
$DOBMonth = $DOBLong.Substring(3,($DOBLong.length - 8)) #from 4th character
$DOBYear = $DOBLong.Substring(($DOBLong.length - 4), 4) #last 4 characters
$DOBMM=$DOBArray.Get_Item($DOBMonth)
$PWDprimary = (($DOBDay)+ ($DOBMM)+ ($DOBYear))
$PWDsecondary = (($DOBDay)+ ($DOBMM)+ ($DOBYear) + 'B!')
$GradClass=$GradClassArr.Get_Item($_.description)
$ID = $_.cn
$Class = $_.Reg
if ($PrimaryYear -match $_.description){
Set-ADAccountPassword -Identity $_.cn -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $PWDprimary -Force) 
Write-Host "Primary account password set:"  $_.cn $_.givenName $_.sn
$body+="Primary account password set: $ID"
}
elseif ($SecondaryYear -match $_.description){
Set-ADAccountPassword -Identity $_.cn -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $PWDsecondary -Force) 
Write-Host "Secondary account password set:"  $_.cn $_.givenName $_.sn
$body+="Secondary account password set: $ID"

}
else {Write-Host "No Year Supplied or no DOB supplied, so no password set:" $_.cn $_.givenName $_.sn
$body+= "No Year Supplied or no DOB supplied, so no password set: $ID"
}
}




#Update fields on the AD MailContact
import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$GradClassArr.Get_Item($_.description)
$Office = ('Graduating Class: ' + $GradClass)
$DOA = $_.DOA
$Today = Get-Date
$Email = ($UserID + '@company.net')
$ID = $_.cn
$DNContact = Get-DistinguishedNameContact $Email
Write-Host $DNContact
Set-Contact -Identity $DNContact -company "British International School, Hanoi" -department $_.department -Office $Office -title $_.title 
Set-ADObject -Identity $DNContact -Description ('GradClass: ' + $GradClass) -Replace @{info=("Date of Admission:" + $DOA + "Updated on:" + $Today)} 
Write-Host "AD Mail Contact fields updated:"  $UserID
$body+="AD Mail Contact fields updated:  $ID"}


#Update fields on the AD User
import-csv “\\mail\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$GradClassArr.Get_Item($_.description)
$HomeDir = ('\\192.168.150.55\studentpro$\' + $UserID)
$Office = ('Graduating Class: ' + $GradClass)
$Email = ($UserID + "@company.net")
$DN = Get-DistinguishedNameUser $UserID
$ID = $_.cn
#Set Grade, Home Drive, Email Address, Grad Class
Set-ADUser -Identity $DN -Description $_.description -Title $_.title  -HomeDirectory $HomeDir -HomeDrive ('S:') -EmailAddress $Email -Office $Office 
#Set House
Set-ADUser -Identity $DN -Department $_.department -Replace @{bishnGender=$_.bishnGender;bishnHouse=$_.bishnHouse;bishnGraduatingClass=$GradClass;info=("Date of Admission:" + $_.DOA + "Updated on:" + $Today)}
#Set Password Settings
Set-ADUser -Identity $DN -PasswordNeverExpires:$true -ChangePasswordAtLogon:$false -enabled:$true
Write-Host "AD User Account fields updated:"   $_.givenName $_.sn $_.cn
$body+="AD User Account fields updated:   $ID"}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@company.com>" -To "Heidi Hendry <heidi.hendry@company.com>" -Subject "Create new AD Student Login Accounts & Mail Contacts" -Body $body -Smtpserver "mail.company.com"
