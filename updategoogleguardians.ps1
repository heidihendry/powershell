cd c:\gam
$body = @()
$peFirst = @()
$peSecond = @()
$peLength = 0
import-csv “C:\scripts\UpdateGoogleGuardians\updategoogleguardians.csv” | ForEach-Object {
$studentemail= $_.studentemail
$parentemail = $_.parentemail
$peLength = $parentemail.length
$peSep1 = $parentemail.IndexOf(";")
$peSep2 = $parentemail.IndexOf(";",$peSep1+1)

if ($peSep1 -gt 0 -And $peSep2 -gt 0) {$peFirst = $parentemail.Substring(0,$peSep1)
$peSecond = $parentemail.Substring($peSep1+1,($peSep2-$peSep1-1))}

else
{$peFirst = $parentemail
$peSecond = @()}


#$peFirst = $parentemail.Substring(0,$peSep1)
#$peSecond = $parentemail.Substring($peSep1+1,($peSep2-$peSep1-1))

$peFirst
$peSecond
$body+=gam create guardianinvite $peFirst $studentemail
$body+=gam create guardianinvite $peSecond $studentemail
Write-Host "Guardian Invite created:"  $studentemail $peFirst $peSecond
$peFirst = @()
$peSecond = @()
$peLength = 0

}
$body = $body | out-string
Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "Heidi Hendry <it-info@bishanoi.com>" -Subject "GAM Update Google Guardians" -Body $body -Smtpserver "mail.bishanoi.com"
