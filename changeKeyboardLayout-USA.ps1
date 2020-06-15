$langList = New-WinUserLanguageList en-au
$langList[0].InputMethodTips.Clear()
$langList[0].InputMethodTips.Add('0C09:00000409') # English (Australia) - Keyboard USA
Set-WinUserLanguageList $langList