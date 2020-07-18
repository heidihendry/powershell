cd c:\gam
$body = @()
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2016-2017 Students\250x250-Google\#user#.jpg" 
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2016-2017 Staff\250x250-Google\#user#.jpg"

$body = $body | out-string

#$email = @{
#From = "GAM - Update Photos <it-info@company.com>"
#To = "Heidi Hendry <heidi.hendry@company.com>" 
#CC = "it-info@company.com"
#Subject = "GAM Update Photos to G-Suite"
#SMTPServer = "mail.company.com"
#Body = $body
#}

Send-MailMessage -From "IT Scripts <it-info@company.com>" -To "IT Info <it-info@company.com>" -Subject "GAM Update Photos to G-Suite 2016-2017" -Body $body -Smtpserver "mail.company.com"

cd c:\gam
$body = @()
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Students\250x250-Google\#user#.jpg" 
$body+= gam all users update photo "\\nas\m\4.Marketing & Graphic Design Materials\6. Photos for IT Databases\2017-2018 Staff\250x250Google\#user#.jpg"

$body = $body | out-string


Send-MailMessage -From "IT Scripts <it-info@company.com>" -To "IT Info <it-info@company.com>" -Subject "GAM Update Photos to G-Suite 2017-2018" -Body $body -Smtpserver "mail.company.com"

exit
