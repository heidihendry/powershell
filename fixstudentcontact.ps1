import-module ActiveDirectory
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
  
# $DN = Get-DistinguishedNameUser $_.mail  
  
Function Get-DistinguishedNameUser ($samAccountName) 
{  
   $searcher = New-Object System.DirectoryServices.DirectorySearcher([ADSI]'') 
   $searcher.Filter = "(&(objectClass=User)(samAccountName=$samAccountName))" 
   $result = $searcher.FindOne() 
 
   Return $result.GetDirectoryEntry().DistinguishedName 
}  



#Create the Mail Contact
import-csv “\\mail\script$\SIMSImport\AllStudents\allstudents.csv” | ForEach-Object {
$Email = ($_.cn + "@COMPANY.net")
New-MailContact -name ($_.givenName + " " + $_.sn + " - " + $_.cn)  -alias $_.cn –FirstName $_.givenName –LastName $_.sn -ExternalEmailAddress $Email -PrimarySmtpAddress $Email -OrganizationalUnit "OU=Student Mail Contacts,OU=Students,OU=AllUsers,OU=BIS-HN,DC=COMPANY,DC=com"  -Displayname ($_.givenName + " " + $_.sn + " - " + $_.cn)
Write-Host "Exchange Mail Contact created:"  $_.cn $_.givenName $_.sn
$body+="Exchange Mail Contact created:  $Email"
}

#Add Mail Contact to Correct Mail Group
import-csv “\\mail\script$\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$Email = ($_.cn + "@COMPANY.net")
$GradClass=$GradClassArr.Get_Item($_.description)
$ID = $_.cn
$Class = $_.Reg
if ($PrimaryYearClassGroups -match $_.description)
{

Add-DistributionGroupMember -Identity ( "COMPANY Class 0" + $_.Reg) -Member $Email
Write-Host "Primary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Primary Account $Email added to class $Class"
 }
elseif ($EYCClass -match $_.description)
{

Add-DistributionGroupMember -Identity ( "COMPANY Class " + $_.Reg) -Member $Email
Write-Host "EYC Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="EYC Account $ID added to class $Class"
}

elseif ($SecondaryKS3 -match $_.description)
{

Add-DistributionGroupMember -Identity ( "COMPANY Class 0" + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}

elseif ($SecondarySenior -match $_.description)
{

Add-DistributionGroupMember -Identity ( "COMPANY Class " + $_.Reg) -Member $Email
Write-Host "Secondary Account" $_.cn $_.givenName $_.sn "added to class" $_.Reg
$body+="Secondary Account $ID added to class $Class"
}
else {

Write-Host "No Class Supplied, so mail contact not added to Class group:" $_.cn $_.givenName $_.sn
$body+="No Class Supplied, so mail contact not added to Class group: $ID" }
}

#Update fields on the AD MailContact
import-csv “\\mail\script$\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$UserID = $_.cn
$GradClass=$GradClassArr.Get_Item($_.description)
$Office = ('Graduating Class: ' + $GradClass)
$DOA = $_.DOA
$Today = Get-Date
$Email = ($UserID + '@COMPANY.net')
$ID = $_.cn
$DNContact = Get-DistinguishedNameContact $UserID
Write-Host $DNContact
Set-Contact -Identity $DNContact -company "British International School, Hanoi" -department $_.department -Office $Office -title $_.title 
Set-ADObject -Identity $DNContact -Description ('GradClass: ' + $GradClass) -Replace @{info=("Date of Admission:" + $DOA + "Updated on:" + $Today)} 
Write-Host "AD Mail Contact fields updated:"  $UserID
$body+="AD Mail Contact fields updated:  $ID"}

$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@COMPANY.com>" -To "Heidi Hendry <heidi.hendry@COMPANY.com>" -Subject "Fix Mail Contacts" -Body $body -Smtpserver "mail.COMPANY.com"
