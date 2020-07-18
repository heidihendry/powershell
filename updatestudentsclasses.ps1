import-module ActiveDirectory
Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010
. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
Connect-ExchangeServer -auto

$PrimaryYear = "Foundation 1","Foundation 2","Foundation 3", "Year 01", "Year 02","Year 03","Year 04","Year 05","Year 06"
$SecondaryYear = "Year 07","Year 08","Year 09","Year 10","Year 11","Year 12","Year 13"
$SecondaryKS3 = "Year 07","Year 08","Year 09"
$SecondarySenior = "Year 10","Year 11","Year 12","Year 13"
$GradClassArr = @{"Year 13" = 2018;"Year 12"=2019;"Year 11"=2020;"Year 10"=2021;"Year 09"=2022;"Year 08"=2023;"Year 07"=2024;"Year 06"=2025;"Year 05"=2026;"Year 04"=2027;"Year 03"=2028;"Year 02"=2029;"Year 01"=2030;"Foundation 3"=2031;"Foundation 2"=2032;"Foundation 1"=2033}
$DOBArray = @{"January" = "01";"February" = "02";"March" = "03";"April"="04";"May"="05";"June"="06";"July"="07";"August"="08";"September"="09";"October"="10";"November"="11";"December"="12"}
$AllClasses = "COMPANY Class 03B","COMPANY Class 03I","COMPANY Class 03S","COMPANY Class 04B","COMPANY Class 04I","COMPANY Class 04S","COMPANY Class 05B","COMPANY Class 05I","COMPANY Class 05S","COMPANY Class 06B","COMPANY Class 06I","COMPANY Class 06S","COMPANY Class 06H","COMPANY Class 07B","COMPANY Class 07I","COMPANY Class 07S","COMPANY Class 08B","COMPANY Class 08I","COMPANY Class 08S","COMPANY Class 09B","COMPANY Class 09H","COMPANY Class 09I","COMPANY Class 09S","COMPANY Class 10B","COMPANY Class 10H","COMPANY Class 10I","COMPANY Class 10S","COMPANY Class 11B","COMPANY Class 11H","COMPANY Class 11I","COMPANY Class 11S","COMPANY Class 12B","COMPANY Class 12I","COMPANY Class 12S","COMPANY Class 13B","COMPANY Class 13I"
$StudentSecurityGroups ="COMPANY_Student_Secondary","COMPANY_Student_Primary"
$body = @()
#Remove the Mail Contact from all Mail Distribution Groups
# Use the removeallclassdistributiongroupmembers.ps1
 
#Add the Mail Contact to the relevant Mail Distribution Group
# what if no class supplied?
import-csv “\\mail\scripts$\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$GradClass=$GradClassArr.Get_Item($_.description)
$Email = ($_.cn + "@COMPANY.com")
$ID = $_.cn
$Class = $_.Reg
if ($PrimaryYear -match $_.description)
{Add-DistributionGroupMember -Identity ( "COMPANY Class 0" + $_.Reg) -Member $Email
Write-Host "Primary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg

$body+="Primary Account $ID added to class $Class"
 }

elseif ($SecondaryKS3 -match $_.description)
{Add-DistributionGroupMember -Identity ( "COMPANY Class 0" + $_.Reg) -Member $Email
Write-HostWrite-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}

elseif ($SecondarySenior -match $_.description)
{Add-DistributionGroupMember -Identity ( "COMPANY Class " + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}
else {Write-Host "No Year Supplied, so no mail contact created:" $_.cn $_.givenName $_.sn
$body+="No Year Supplied, so no mail contact created: $ID "
}
}

#Add the AD User to the relevant Security Group

import-csv “\\mail\scripts$\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {

$GradClass=$GradClassArr.Get_Item($_.description)
if ($PrimaryYear -contains $_.description)
{
$ID = $_.cn	

$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=COMPANY_Student_Primary,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=COMPANY,DC=com" –Member $DN
Write-Host "Primary account updated:"  $_.cn $_.givenName $_.sn
$body+= "Primary account updated: $ID"
}
elseif ($SecondaryYear -contains $_.description)
{
$ID = $_.cn

$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=COMPANY_Student_Secopndary,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=COMPANY,DC=com" –Member $DN
Write-Host "Secondary account updated:"  $_.cn $_.givenName $_.sn
$body+="Secondary account updated: $ID"
}
else {
$ID = $_.cn
Write-Host "No Year Supplied, so no user created:" $_.cn $_.givenName $_.sn
$body+="No Year Supplied, so no user created: $ID"
}
}



# Get the user's Distinguished Name 
Function Get-DistinguishedNameContact ($strMail) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=Contact)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
} 
  
Function Get-DistinguishedNameUser ($strMail) #$samAccountName ??
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(mail=$strMail))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}   



#Update fields on the AD MailContact
import-csv “\\mail\scripts$\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$GradClass=$GradClassArr.Get_Item($_.description)
$Office = ('Graduating Class: ' + $GradClass)
$DOA = $_.DOA
$Today = Get-Date
$Email = ($_.cn + '@COMPANY.com')
$DN = Get-DistinguishedNameContact $Email
Set-Contact -Identity $DN -company "British International School, Hanoi" -department $_.department -Office $Office -title $_.title 
Set-ADObject -Identity $DN -Description ('GradClass: ' + $GradClass) # -Replace @{info=("Date of Admission:" + $DOA + "Updated on:" + $Today)} 
Write-Host "MailContact fields updated:"  $DN
$body+="MailContact fields updated:  $Email"
}


#Update fields on the AD User
import-csv “\\mail\scripts$\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$GradClassArr.Get_Item($_.description)
$Email = ($_.cn + "@COMPANY.com")
$HomeDir = ('\\nas.COMPANY.com\studentpro$\' + $UserID)
$Office = ('Graduating Class: ' + $GradClass)
$DN = Get-DistinguishedNameUser $Email

Set-ADUser -Identity $DN -Description $_.description -Title $_.title -Department $_.department -HomeDirectory $HomeDir -HomeDrive ('S:') -EmailAddress $Email -Office $Office -Replace @{COMPANYGender=$_.COMPANYGender;COMPANYHouse=$_.COMPANYHouse;COMPANYGraduatingClass=$GradClass;info=("Date of Admission:" + $_.DOA + "Updated on:" + $Today)} -PasswordNeverExpires:$true -ChangePasswordAtLogon:$false -enabled:$true
Write-HostWrite-Host "Account fields updated:"  $_.cn $_.givenName $_.sn
$body+="Account fields updated: $UserID" 
}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@COMPANY.com>" -To "Heidi Hendry <heidi.hendry@COMPANY.com>" -Subject "Update Student Classes" -Body $body -Smtpserver "mail.COMPANY.com"

# Move all Year 6 to Secondary
# Check users are not in Leavers
