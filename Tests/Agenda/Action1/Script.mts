
' Définir la variable pour le titre du rendez-vous
Dim titreRdv

' Assigner une valeur à la variable
titreRdv = "uft"

' Accéder à Gmail
Wait(2) ' Attendre que la page est complètement chargée
Browser("Google").Page("Google").Link("Gmail").Click

' Sélectionner l'onglet ( Agenda)
Browser("Google").Page("Boîte de réception (3)").WebTabStrip("WebTabStrip").Select "#0"

Wait(5) ' Attendre que la page soit chargée
' Cliquer sur l'heure "12 PM" pour ajouter un rendez-vous
Browser("Google").Page("Boîte de réception (3)").Frame("I0_1725372715061").WebElement("12 PM").Click


' Ajouter un titre au rendez-vous
Browser("Boîte de réception (3)").Page("Boîte de réception (3)").Frame("I0_1725376856498").WebEdit("Ajouter un titre").Set titreRdv

' Enregistrer le rendez-vous
Browser("Boîte de réception (3)").Page("Boîte de réception (3)").Frame("I0_1725376856498").WebButton("Enregistrer").Click

' Vérification via un checkpoint
Window("Window").WinObject("Applications en cours").Check CheckPoint("Applications en cours d’exécution")

' Fermer la fenêtre de l'événement
Browser("Google").Page("Boîte de réception (3)").Frame("I0_1725372715061").WebButton("Fermer").Click
 @@ script infofile_;_ZIP::ssf11.xml_;_
