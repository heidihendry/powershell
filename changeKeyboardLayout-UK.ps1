$langList = New-WinUserLanguageList en-GB
$langList[0].InputMethodTips.Clear()
$langList[0].InputMethodTips.Add('0809:00000809') # English (UK) - Keyboard (UK)
Set-WinUserLanguageList $langList