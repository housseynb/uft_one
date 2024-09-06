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
