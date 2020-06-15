import-module ActiveDirectory


import-csv “c:\scripts\ExpiredPasswords\passwordsLastSet2.csv” | ForEach-Object {
$User = '[ADSI]"LDAP://' + $_.DistinguishedName + '"'
$User.pwdLastSet = 0
$User.SetInfo()
Write-Host $User
}