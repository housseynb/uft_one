' Action de Déconnexion (Logout)

' Cliquer sur le bouton de compte Google
Browser("Boîte de réception (2)").Page("Boîte de réception (2)").WebButton("Compte Google houss uft").Click
' Attendre que le menu de compte soit visible
Browser("Boîte de réception (2)").Page("Boîte de réception (2)").WaitProperty "visible", True, 10

' Cliquer sur "Se déconnecter"
Browser("Boîte de réception (2)").Page("Boîte de réception (2)").Frame("account").Link("Se déconnecter").Click
' Attendre que la page de confirmation de déconnexion soit visible
Browser("Boîte de réception (2)").Page("Gmail").WaitProperty "visible", True, 10

' Cliquer sur "Supprimer un compte"
Browser("Boîte de réception (2)").Page("Gmail").Link("Supprimer un compte").Click
' Attendre que la liste des comptes soit visible
Browser("Boîte de réception (2)").Page("Gmail").WaitProperty "visible", True, 10

' Cliquer sur le compte à supprimer
Browser("Boîte de réception (2)").Page("Gmail").Link("houss ufthoussuft@gmail.comDéc").Click
' Attendre que le bouton de confirmation soit visible
Browser("Boîte de réception (2)").Page("Gmail").WaitProperty "visible", True, 10

' Cliquer sur "Oui, supprimer"
Browser("Boîte de réception (2)").Page("Gmail").WebButton("Oui, supprimer").Click
' Attendre la confirmation de la suppression
Wait(5) @@ script infofile_;_ZIP::ssf5.xml_;_
