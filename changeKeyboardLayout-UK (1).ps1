$langList = New-WinUserLanguageList en-au
$langList[0].InputMethodTips.Clear()
$langList[0].InputMethodTips.Add('0C09:00000809') # English (Australia) - Keyboard Great Britain
Set-WinUserLanguageList $langList