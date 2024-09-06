' Initialisation du navigateur et de la page
Set browserObj = Browser("JPetStore Demo")
Set home = browserObj.Page("Home")

' Cliquer sur le lien "Sign Out"
home.Link("Sign Out").Click

' Synchroniser la visibilité de lien "Sign In"
home.Link("Sign In").WaitProperty "visible", True, 5000
