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
$AllClasses = "BISHN Class 03B","BISHN Class 03I","BISHN Class 03S","BISHN Class 04B","BISHN Class 04I","BISHN Class 04S","BISHN Class 05B","BISHN Class 05I","BISHN Class 05S","BISHN Class 06B","BISHN Class 06I","BISHN Class 06S","BISHN Class 06H","BISHN Class 07B","BISHN Class 07I","BISHN Class 07S","BISHN Class 08B","BISHN Class 08I","BISHN Class 08S","BISHN Class 09B","BISHN Class 09H","BISHN Class 09I","BISHN Class 09S","BISHN Class 10B","BISHN Class 10H","BISHN Class 10I","BISHN Class 10S","BISHN Class 11B","BISHN Class 11H","BISHN Class 11I","BISHN Class 11S","BISHN Class 12B","BISHN Class 12I","BISHN Class 12S","BISHN Class 13B","BISHN Class 13I"
$StudentSecurityGroups ="BISHN_Student_Secondary","BISHN_Student_Primary"
$body = @()
#Remove the Mail Contact from all Mail Distribution Groups
# Use the removeallclassdistributiongroupmembers.ps1
 
#Add the Mail Contact to the relevant Mail Distribution Group
# what if no class supplied?
import-csv “\\mail\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$GradClass=$GradClassArr.Get_Item($_.description)
$Email = ($_.cn + "@bishanoi.com")
$ID = $_.cn
$Class = $_.Reg
if ($PrimaryYear -match $_.description)
{Add-DistributionGroupMember -Identity ( "BISHN Class 0" + $_.Reg) -Member $Email
Write-Host "Primary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg

$body+="Primary Account $ID added to class $Class"
 }

elseif ($SecondaryKS3 -match $_.description)
{Add-DistributionGroupMember -Identity ( "BISHN Class 0" + $_.Reg) -Member $Email
Write-HostWrite-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}

elseif ($SecondarySenior -match $_.description)
{Add-DistributionGroupMember -Identity ( "BISHN Class " + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}
else {Write-Host "No Year Supplied, so no mail contact created:" $_.cn $_.givenName $_.sn
$body+="No Year Supplied, so no mail contact created: $ID "
}
}

#Add the AD User to the relevant Security Group

import-csv “\\mail\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {

$GradClass=$GradClassArr.Get_Item($_.description)
if ($PrimaryYear -contains $_.description)
{
$ID = $_.cn	

$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=BISHN_Student_Primary,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=bishanoi,DC=com" –Member $DN
Write-Host "Primary account updated:"  $_.cn $_.givenName $_.sn
$body+= "Primary account updated: $ID"
}
elseif ($SecondaryYear -contains $_.description)
{
$ID = $_.cn

$DN = Get-DistinguishedNameUser $ID
Add-ADGroupMember "CN=BISHN_Student_Secopndary,OU=Security Groups,OU=Groups,OU=BIS-HN,DC=bishanoi,DC=com" –Member $DN
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
import-csv “\\mail\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$GradClass=$GradClassArr.Get_Item($_.description)
$Office = ('Graduating Class: ' + $GradClass)
$DOA = $_.DOA
$Today = Get-Date
$Email = ($_.cn + '@bishanoi.com')
$DN = Get-DistinguishedNameContact $Email
Set-Contact -Identity $DN -company "British International School, Hanoi" -department $_.department -Office $Office -title $_.title 
Set-ADObject -Identity $DN -Description ('GradClass: ' + $GradClass) # -Replace @{info=("Date of Admission:" + $DOA + "Updated on:" + $Today)} 
Write-Host "MailContact fields updated:"  $DN
$body+="MailContact fields updated:  $Email"
}


#Update fields on the AD User
import-csv “\\mail\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$GradClassArr.Get_Item($_.description)
$Email = ($_.cn + "@bishanoi.com")
$HomeDir = ('\\nas.bishanoi.com\studentpro$\' + $UserID)
$Office = ('Graduating Class: ' + $GradClass)
$DN = Get-DistinguishedNameUser $Email

Set-ADUser -Identity $DN -Description $_.description -Title $_.title -Department $_.department -HomeDirectory $HomeDir -HomeDrive ('S:') -EmailAddress $Email -Office $Office -Replace @{bishnGender=$_.bishnGender;bishnHouse=$_.bishnHouse;bishnGraduatingClass=$GradClass;info=("Date of Admission:" + $_.DOA + "Updated on:" + $Today)} -PasswordNeverExpires:$true -ChangePasswordAtLogon:$false -enabled:$true
Write-HostWrite-Host "Account fields updated:"  $_.cn $_.givenName $_.sn
$body+="Account fields updated: $UserID" 
}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "Heidi Hendry <heidi.hendry@bishanoi.com>" -Subject "Update Student Classes" -Body $body -Smtpserver "mail.bishanoi.com"

# Move all Year 6 to Secondary
# Check users are not in Leavers