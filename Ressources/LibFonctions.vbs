' Fonction pour vérifier l'existence d'un élément
Function VerifElement(Element, nomElement, Etat)
    Dim existence
    existence = Element.Exist(2)
    
    Select Case Etat
        Case "Passant"
            If existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Elément", "L'élément " & nomElement & " a été trouvé"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Elément", "L'élément " & nomElement & " n'a pas été trouvé: Le test est en échec: KO"
            End If
        Case "NonPassant"
            If Not existence Then
                VerifElement = "OK"
                Reporter.ReportEvent micPass, "Verif Elément", "L'élément " & nomElement & " n'a pas été trouvé comme attendu"
            Else
                VerifElement = "KO"
                Reporter.ReportEvent micFail, "Verif Elément", "L'élément " & nomElement & " a été trouvé contrairement à l'attendu"
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
        Reporter.ReportEvent micFail, "Synchronisation", "L'élément n'est pas devenu visible après " & attenteMax & " secondes"
    End If
End Sub

' Fonction pour vérifier la propriété innertext d'un élément
Function VerifierInnerText(element, texteAttendu)
    Dim innerText
    innerText = element.GetROProperty("innertext")
    
    If innerText = texteAttendu Then
        VerifierInnerText = "OK"
        Reporter.ReportEvent micPass, "Vérification InnerText", "L'innertext de l'élément est correct : " & innerText
    Else
        VerifierInnerText = "KO"
        Reporter.ReportEvent micFail, "Vérification InnerText", "L'innertext de l'élément est incorrect : " & innerText
    End If
End Function
