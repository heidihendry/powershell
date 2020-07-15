#==================| Satnaam Waheguru Ji |===============================   
#              
#            Author  :  Aman Dhally    
#            E-Mail  :  amandhally@gmail.com    
#            website :  www.amandhally.net    
#            twitter :   @AmanDhally    
#            blog    : http://newdelhipowershellusergroup.blogspot.in/   
#            facebook: http://www.facebook.com/groups/254997707860848/    
#            Linkedin: http://www.linkedin.com/profile/view?id=23651495    
#    
#            Creation Date    : 09-12-2013 
#            File    :           
#            Purpose :        
#            Version : 1    
#            
#   
#           
#========================================================================   
 
##Note ====> Before running this script, make sure you have RSAT tool installed. 
 
#Import Module Active Directory 
Import-Module ActiveDirectory -ErrorAction 'Stop' 
 
# Days after password expires, Change the days as per your Default Paaaword Expiration group Policy 
[int]$totalDays = 90 
 
# TOday 
$todayDate =  Get-Date 
 
 
#Password expiredCollection 
$passwordExpiredCollection = @() 
 
# Email Option and Value  
 
$smtp = "mail.company.com" 
$subject = "Please change your password soon" 
 
# filtering user from AD 
$adUsers = Get-ADUser -Filter {(ObjectClass -eq "user") -and (EmailAddress -ne "$null")  -and (PasswordNeverExpires -eq "False") -and (Enabled -eq $true) } -Properties PasswordNeverExpires,PasswordLastSet,PasswordExpired,LockedOut,EmailAddress 
 
foreach ( $aduser in $adUsers) 
 
        { 
     
           if ($aduser.PasswordLastSet -ne $null) {  
 
             
            [datetime]$lastPasswordSet = $aduser.PasswordLastSet 
            $timeSpan = New-TimeSpan  (Get-date -Date $lastPasswordSet.Date ) 
            $expirationTime = $totalDays - $timeSpan.Days 
            
            } 
 
 
            Switch ($expirationTime) 
            { 
 
 
            7  { 
                    $dateAfter7Days = (Get-Date).AddDays(7).ToShortDateString().ToString() 
                       $passwordExpiring7Days  += $aduser.Name + ";" + $aduser.EmailAddress + ";" + $expirationTime + ";" + $dateAfter7Days 
             
                } 
             
 
             
             
            } 
 
            #switch stop 
 
 
            # If User password is expired. 
 
            if ( $aduser.PasswordExpired -eq $true )  
                 
                { 
             
                    $passwordExpiredCollection += $aduser.Name + ";" + $aduser.EmailAddress + ";" + $expirationTime + "`n" 
             
                } 
 
 
 
         
        } 
 
 
 
# Splitting 
 
 
if ( $passwordExpiring7Days -ne $null ) { 
 
        foreach ( $7name in $passwordExpiring7Days  ) { 
 
 
            $7userCollection = $7name -split ";" 
            $7userName = $7userCollection[0] 
            $7userEmail = $7userCollection[1] 
            $7pass = $7userCollection[2] 
            $7day = $7userCollection[3] 
 
 
            Write-Host "Dear $7userName, Confirming your email address is $7userEmail. Please be aware that your Company password is expiring in $7pass days." -ForegroundColor Green 
 
            $body = "Dear $7userName, <br>" 
             
            $body += "<br>" 
            $body += "Please be aware that your Company password is expiring in  <b><font color=red> $7pass days</b></font>. Please ensure you have changed it before then.<br>" 
			 $body += "Your password can be changed on your Company computer by using Ctrl-Alt-Del and then 'Change Password'. This will automatically change the password for all your other Company services.<br>" 
            $body += "<br>" 
 
            $body += "Regards<br>" 
            $body += "I.T. Team<br>" 
            $body += "<br>" 
            $body += "<br>" 
            $body += "<b>How to change your password:</b><br>" 
            $body += "    1. Press CTRL+ALT+DELETE, and then click Change a password.<br>" 
            $body += "    2. Type your old password, type your new password, type your new password again to confirm it, and then press ENTER.<br>" 
 
            # if you want to send an email, please un-comment the below line. 
            Send-MailMessage -to $7userEmail -From "helpdesk@company.com"  -SmtpServer $smtp -Body $body -BodyAsHtml -Subject $subject  -Priority high -Encoding UTF8 
             
             
         
            } 
 
} 
 
 
# sending list of password expired. 
  
 $body = "" 
 $body += $passwordExpiredCollection 
 
 Write-Warning "Users whose passwords are already expired ========"     
 Write-Host $passwordExpiredCollection     
 
# if you want to send an email, please un-comment the below line. 
Send-MailMessage -to "it-info@company.com" -SmtpServer $smtp -From "it-info@company.com" -Body $body -Subject "Passwords expired" 
