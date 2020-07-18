$PrimaryYear = "Foundation 1","Foundation 2","Foundation 3", "Year 01", "Year 02","Year 03","Year 04","Year 05","Year 06"
$EYCCLass = "Foundation 1","Foundation 2","Foundation 3"
$PrimaryYearClassGroups = "Year 01", "Year 02","Year 03","Year 04","Year 05","Year 06"
$SecondaryYear = "Year 07","Year 08","Year 09","Year 10","Year 11","Year 12","Year 13"
$SecondaryKS3 = "Year 07","Year 08","Year 09"
$SecondarySenior = "Year 10","Year 11","Year 12","Year 13"
$GradClassArr = @{"Year 13" = 2018;"Year 12"=2019;"Year 11"=2020;"Year 10"=2021;"Year 09"=2022;"Year 08"=2023;"Year 07"=2024;"Year 06"=2025;"Year 05"=2026;"Year 04"=2027;"Year 03"=2028;"Year 02"=2029;"Year 01"=2030;"Foundation 3"=2031;"Foundation 2"=2032;"Foundation 1"=2033}
$DOBArray = @{"January" = "01";"February" = "02";"March" = "03";"April"="04";"May"="05";"June"="06";"July"="07";"August"="08";"September"="09";"October"="10";"November"="11";"December"="12"}
$body = @()

#Set Password 
import-csv “\\mail\scripts$\SIMSImport\NewStudents\newstudents.csv” | ForEach-Object {
$DOBLong = $_.DOB
$DOBDay = $DOBLong.Substring(0,2) #first 2 characters
$DOBMonth = $DOBLong.Substring(3,($DOBLong.length - 8)) #from 4th character
$DOBYear = $DOBLong.Substring(($DOBLong.length - 4), 4) #last 4 characters
$DOBMM=$DOBArray.Get_Item($DOBMonth)
$PWDprimary = (($DOBDay)+ ($DOBMM)+ ($DOBYear))
$PWDsecondary = (($DOBDay)+ ($DOBMM)+ ($DOBYear) + 'B!')
$GradClass=$GradClassArr.Get_Item($_.description)
$ID = ($_.cn + "@COMPANY.net")
$Class = $_.Reg
if ($PrimaryYear -match $_.description){
gam update user $ID password $PWDprimary changepassword off
Write-Host "Primary account password set:"  $ID $_.cn $_.givenName $_.sn $PWDPrimary
$body+="Primary account password set: $ID "
}
elseif ($SecondaryYear -match $_.description){
gam update user $ID password $PWDsecondary changepassword off
Write-Host "Secondary account password set:" $ID  $_.cn $_.givenName $_.sn $PWDSecondary
$body+="Secondary account password set: $ID "

}
else {Write-Host "No Year Supplied or no DOB supplied, so no password set:" $_.cn $_.givenName $_.sn
$body+= "No Year Supplied or no DOB supplied, so no password set: $ID"
}
}


$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@COMPANY.com>" -To "IT Info <it-info@COMPANY.com>" -Subject "New Student Google Password Reset" -Body $body -Smtpserver "mail.COMPANY.com"
