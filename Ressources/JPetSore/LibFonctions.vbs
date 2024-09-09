' Fonction pour verifier l'existence d'un element
Function VerifElement(Element, nomElement, Etat)
    Dim existence
    existence = Element.Exist(2)
    
    Select Case Etat
        Case "Passant"
            If existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Element", "L'element " & nomElement & " a ete trouve"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Element", "L'element " & nomElement & " n'a pas ete trouve: Le test est en echec: KO"
            End If
        Case "NonPassant"
            If Not existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Element", "L'element " & nomElement & " n'a pas ete trouve comme attendu"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Element", "L'element " & nomElement & " a ete trouve contrairement a l'attendu"
            End If
    End Select
End Function

' Procedure pour synchroniser la visibilite d'un element
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
        Reporter.ReportEvent micFail, "Synchronisation", "L'element n'est pas visible apres " & attenteMax & " secondes"
    End If
End Sub

' Fonction pour verifier la propriete innertext d'un element
Function VerifierInnerText(element, texteAttendu)
    Dim innerText
    innerText = Trim(element.GetROProperty("innertext"))
    
    If innerText = texteAttendu Then
        VerifierInnerText = "OK"
        Reporter.ReportEvent micPass, "Verification InnerText", "L'innertext de l'element est correct : " & innerText
    Else
        VerifierInnerText = "KO"
        Reporter.ReportEvent micFail, "Verification InnerText", "L'innertext de l'element est incorrect : " & innerText
    End If
End Function

' Fonction  pour vérifier l'état d'une case à cocher
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

    ' Obtenir la valeur actuelle de la propriété "value"
    valeurActuelle = element.GetROProperty("value")
    
    ' Comparer la valeur actuelle avec la valeur attendue
    If CStr(valeurActuelle) = CStr(valeurAttendue) Then
        VerifierValeurElement = "OK"
        Reporter.ReportEvent micPass, "Verification Valeur Element", "La valeur de l'élément est correcte : " & valeurActuelle
    Else
        VerifierValeurElement = "KO"
        Reporter.ReportEvent micFail, "Verification Valeur Element", "La valeur de l'élément est incorrecte : " & valeurActuelle
    End If
End Function


