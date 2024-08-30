Dim url, chromePath
Dim WshShell
Dim ePageAccueil
Dim ePageLogin
Set ePageAccueil = Browser("Google").Page("Advantage Shopping")
Set WshShell = CreateObject("WScript.Shell")

'définir url + chemin chrome
url = "https://www.advantageonlineshopping.com/#/"  
chromePath = "C:\Program Files\Google\Chrome\Application\chrome.exe" 

'ouvrir chrome sur Advantage Shopping
WshShell.Run """" & chromePath & """ -url " & url

'Étape 1 : Cliquer sur le bouton login
ePageAccueil.Sync
ePageAccueil.Link("UserMenu").Click
ePageAccueil.WebEdit("username").Set "houboudje"
ePageAccueil.WebEdit("password").SetSecure "66d17a54657d0bb64386607e7a4e20edd72e581a28705bbc"
ePageAccueil.WebButton("sign_in_btn").Click
ePageAccueil.Sync @@ script infofile_;_ZIP::ssf66.xml_;_
