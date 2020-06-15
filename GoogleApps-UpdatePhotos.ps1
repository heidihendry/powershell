cd c:\gam
$body = @()
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2016-2017 Students\250x250-Google\#user#.jpg" 
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2016-2017 Staff\250x250-Google\#user#.jpg"

$body = $body | out-string

#$email = @{
#From = "GAM - Update Photos <it-info@bishanoi.com>"
#To = "Heidi Hendry <heidi.hendry@bishanoi.com>" 
#CC = "it-info@bishanoi.com"
#Subject = "GAM Update Photos to G-Suite"
#SMTPServer = "mail.bishanoi.com"
#Body = $body
#}

Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "IT Info <it-info@bishanoi.com>" -Subject "GAM Update Photos to G-Suite 2016-2017" -Body $body -Smtpserver "mail.bishanoi.com"

cd c:\gam
$body = @()
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Students\250x250-Google\#user#.jpg" 
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Staff\250x250Google\#user#.jpg"

$body = $body | out-string


Send-MailMessage -From "IT Scripts <it-info@bishanoi.com>" -To "IT Info <it-info@bishanoi.com>" -Subject "GAM Update Photos to G-Suite 2017-2018" -Body $body -Smtpserver "mail.bishanoi.com"

exit