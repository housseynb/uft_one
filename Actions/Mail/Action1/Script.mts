' Définir les variables pour l'objet du mail et le contenu du mail
Dim adresseEmail, objetEmail, corpsEmail

' Assigner des valeurs aux variables
adresseEmail = "houssuft@gmail.com"
objetEmail = "Test mail"
corpsEmail = "Xavier is great than Mister J"

' Accéder à Gmail
Wait(5) ' Attendre que la page est complètement chargée
Browser("Google").Page("Google").Link("Gmail").Click

' Attendre que la page de la boîte de réception se charge
Browser("Google").Page("Boîte de réception - houssuft@").WaitProperty "visible", True, 10
Reporter.ReportEvent micDone, "Accès Gmail", "Page de la boîte de réception ouverte."

' Composer et envoyer le mail
Browser("Google").Page("Boîte de réception - houssuft@").WebButton("Nouveau message").Click
Browser("Google").Page("Boîte de réception - houssuft@").WaitProperty "visible", True, 10

Browser("Google").Page("Boîte de réception - houssuft@").WebEdit("Destinataires").Set adresseEmail
Browser("Google").Page("Boîte de réception - houssuft@").WebEdit("Objet").Set objetEmail
Browser("Google").Page("Boîte de réception - houssuft@").WebEdit("Corps du message").Set corpsEmail
Browser("Google").Page("Boîte de réception - houssuft@").WebButton("Envoyer ‪(Ctrl-Enter)‬").Click

' Attendre que le mail soit envoyé 
Wait(5)

' Ouvrir le mail envoyé
Browser("Google").Page("Boîte de réception - houssuft@").Link("Boîte de réception 1 message").Click
Browser("Google").Page("Boîte de réception - houssuft@").WebElement(":bf").WaitProperty "visible", True, 10

' Capturer le contenu du mail
Dim texteMessage
texteMessage = Browser("Google").Page("Boîte de réception - houssuft@").WebElement(":bf").GetROProperty("innertext")
Reporter.ReportEvent micDone, "Texte du message", "Texte capturé du message : " & texteMessage

' Vérifier le contenu du mail
If InStr(texteMessage, corpsEmail) > 0 Then
    Reporter.ReportEvent micDone, "Vérification du message", "Le message contient le texte attendu : " & corpsEmail
Else
    Reporter.ReportEvent micFail, "Vérification du message", "Le message ne contient pas le texte attendu : " & corpsEmail
End If
