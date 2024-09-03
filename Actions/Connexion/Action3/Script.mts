' Cliquer sur le bouton de compte Google
Browser("Boîte de réception (2)").Page("Boîte de réception (2)").WebButton("Compte Google houss uft").Click

' Cliquer sur "Se déconnecter"
Browser("Boîte de réception (2)").Page("Boîte de réception (2)").Frame("account").Link("Se déconnecter").Click

' Cliquer sur "Supprimer un compte"
Browser("Boîte de réception (2)").Page("Gmail").Link("Supprimer un compte").Click

' Cliquer sur le compte à supprimer
Browser("Boîte de réception (2)").Page("Gmail").Link("houss ufthoussuft@gmail.comDéc").Click

' Cliquer sur "Oui, supprimer"
Browser("Boîte de réception (2)").Page("Gmail").WebButton("Oui, supprimer").Click
' Attendre la confirmation de la suppression
Wait(3) @@ script infofile_;_ZIP::ssf5.xml_;_
