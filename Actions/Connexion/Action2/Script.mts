' Variables pour l'email et le mot de passe
Dim email, password

' Assigner des valeurs aux variables
email = "houssuft@gmail.com"
password = "66d6d14e267761198a35e9c4ef59726149e686d46ea5bd8976c20bb8"

' Assurer que la page d'accueil est complètement chargée
Browser("nav").Page("home").WaitProperty "visible", True, 10

' Actions de connexion sur le site Google
Browser("nav").Page("home").Link("Sign in").Click

Browser("nav").Page("Sign in - Google Accounts").WebEdit("Email or phone").Set email
Browser("nav").Page("Sign in - Google Accounts").WebButton("Next").Click

Browser("nav").Page("Sign in - Google Accounts").WebEdit("Enter your password").SetSecure password
Browser("nav").Page("Sign in - Google Accounts").WebButton("Next").Click @@ script infofile_;_ZIP::ssf5.xml_;_
