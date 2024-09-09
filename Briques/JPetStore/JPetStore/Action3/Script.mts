' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set home = browserObj.Page("Home")
Set login = browserObj.Page("Login")

' Cliquer sur le lien "Sign In"
home.Link("Sign In").Click

' Saisie des identifiants
login.WebEdit("username").Set Parameter("user")  ' Utilisation du paramètre user

' Utilisation de SetSecure pour le mot de passe
login.WebEdit("password").SetSecure Parameter("password")  ' Utilisation sécurisée du paramètre password

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
