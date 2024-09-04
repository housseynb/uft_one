' Définir l'URL et le chemin de Mozilla Firefox
url = "https://www.google.fr/"
'firefoxPath = "C:\Program Files\Mozilla Firefox\firefox.exe"

SystemUtil.Run "firefox.exe", url

'' Créer une instance de l'objet WshShell
'Set WshShell = CreateObject("WScript.Shell")
'
'' Ouvrir Mozilla Firefox sur l'URL 
'WshShell.Run """" & firefoxPath & """ -url " & url
'
'' Libérer l'objet WshShell
'Set WshShell = Nothing
'
' Attendre que le navigateur se lance
Wait(3)
