' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set home = browserObj.Page("Home")
Set login = browserObj.Page("Login")

' Définition des variables utilisateur et mot de passe
Dim username, password
username = "j2ee"
password = "j2ee"

' Cliquer sur le lien "Sign In"
home.Link("Sign In").Click

' Saisie des identifiants
login.WebEdit("username").Set username
login.WebEdit("password").Set password

' Cliquer sur le bouton "Login"
login.WebButton("Login").Click

' Synchroniser la visibilité de l'élément "WelcomeContent"
SynchroniserVisibilite home.WebElement("WelcomeContent"), 10

' Vérification de la propriété innertext de l'élément "WelcomeContent"
Dim resultatVerification
resultatVerification = VerifierInnerText(home.WebElement("WelcomeContent"), "Welcome ABC!")

' Vérification finale du test
If resultatVerification = "OK" Then
    Reporter.ReportEvent micPass, "Test de connexion", "Le test de connexion est réussi"
Else
    Reporter.ReportEvent micFail, "Test de connexion", "Le test de connexion a échoué"
End If

