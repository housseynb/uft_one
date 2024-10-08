' Fonction pour vérifier l'existence d'un élément
Function VerifElement(Element, nomElement, Etat)
    Dim existence
    existence = Element.Exist(2)
    
    Select Case Etat
        Case "Passant"
            If existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Element", "L'élément " & nomElement & " a été trouvé"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Element", "L'élément " & nomElement & " n'a pas été trouvé: Le test est en échec: KO"
            End If
        Case "NonPassant"
            If Not existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Element", "L'élément " & nomElement & " n'a pas été trouvé comme attendu"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Element", "L'élément " & nomElement & " a été trouvé contrairement à l'attendu"
            End If
    End Select
End Function

' Procédure pour synchroniser la visibilité d'un élément
Sub SynchroniserVisibilite(element, attenteMax)
    Dim cpt
    cpt = 0
    Do
        element.RefreshObject
        If element.GetROProperty("visible") = True Then
            Exit Do
        End If
        cpt = cpt + 1
        Wait(1) ' Attendre 1 seconde
    Loop Until cpt >= attenteMax

    If cpt >= attenteMax Then
        Reporter.ReportEvent micFail, "Synchronisation", "L'élément n'est pas visible après " & attenteMax & " secondes"
    End If
End Sub

' Fonction pour vérifier la propriété innertext d'un élément
Function VerifierInnerText(element, texteAttendu)
    Dim innerText
    innerText = Trim(element.GetROProperty("innertext"))
    
    If innerText = texteAttendu Then
        VerifierInnerText = "OK"
        Reporter.ReportEvent micPass, "Verification InnerText", "L'innertext de l'élément est correct : " & innerText
    Else
        VerifierInnerText = "KO"
        Reporter.ReportEvent micFail, "Verification InnerText", "L'innertext de l'élément est incorrect : " & innerText
    End If
End Function

' Fonction pour vérifier l'état d'une case à cocher
Function VerifierCheckBox(checkbox, etatAttendu)
    Dim etat
    etat = checkbox.GetROProperty("checked") ' Obtenir la valeur de la propriété "checked"
    
    ' Comparer l'état actuel avec l'état attendu
    If etat = etatAttendu Then
        Reporter.ReportEvent micPass, "Verification Case à Cocher", "Case à cocher correcte"
    Else
        Reporter.ReportEvent micFail, "Verification Case à Cocher", "Case à cocher incorrecte"
    End If
End Function

' Fonction pour vérifier la propriété "value" d'un élément
Function VerifierValeur(element, valeurAttendue)
    Dim valeurActuelle
    valeurActuelle = element.GetROProperty("value")
    
    ' Comparer la valeur actuelle avec la valeur attendue
    If CStr(valeurActuelle) = CStr(valeurAttendue) Then
        VerifierValeur = "OK"
        Reporter.ReportEvent micPass, "Verification Valeur Element", "La valeur de l'élément est correcte : " & valeurActuelle
    Else
        VerifierValeur = "KO"
        Reporter.ReportEvent micFail, "Verification Valeur Element", "La valeur de l'élément est incorrecte : " & valeurActuelle
    End If
End Function

' Fonction pour extraire la valeur numérique d'une chaîne de texte contenant un montant
Function ExtraireValeur(texte)
    Dim montant
    
    ' Afficher la valeur pour le débogage
    Reporter.ReportEvent micInfo, "Valeur Récupérée", "Valeur extraite : " & texte
    
    ' Enlever le symbole $ et les espaces
    texte = Trim(Replace(texte, "$", ""))
    
    ' Convertir en nombre
    montant = ConvertirEnFloat(texte)
    
    ' Vérifier si la conversion a échoué
    If montant = -1 Then
        Reporter.ReportEvent micFail, "Erreur de Conversion", "Impossible de convertir la valeur : " & texte
    Else
        ' Afficher la valeur convertie pour le débogage
        Reporter.ReportEvent micInfo, "Conversion Réussie", "Valeur convertie : " & montant
    End If
    
    ' Retourner la valeur extraite
    ExtraireValeur = montant
End Function

Function ConvertirEnFloat(strNombre)
    Dim resultat
    
    ' Afficher la valeur d'entrée
    Reporter.ReportEvent micInfo, "Conversion", "Tentative de conversion de : " & strNombre
    
    ' Remplacer le point par une virgule
    strNombre = Replace(strNombre, ".", ",")
    
    ' Tentative de conversion
    On Error Resume Next
    resultat = CDbl(strNombre)
    If Err.Number = 0 Then
        Reporter.ReportEvent micInfo, "Conversion Réussie", "Résultat : " & resultat
        ConvertirEnFloat = resultat
    Else
        ' Si la première tentative échoue, essayons avec le point décimal
        Err.Clear
        strNombre = Replace(strNombre, ",", ".")
        resultat = CDbl(strNombre)
        If Err.Number = 0 Then
            Reporter.ReportEvent micInfo, "Conversion Réussie", "Résultat : " & resultat
            ConvertirEnFloat = resultat
        Else
            Reporter.ReportEvent micWarning, "Erreur de Conversion", "Erreur lors de la conversion : " & Err.Description
            ConvertirEnFloat = -1
        End If
    End If
    On Error GoTo 0
End Function
