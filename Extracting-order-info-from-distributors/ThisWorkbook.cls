VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

' =============================================================================
' Macros VBA de tout le classeur "ThisWorkbook"
' -----------------------------------------------------------------------------
' ©2021 Arnauld Biganzoli, CNRS-UMR5213-LAPLACE, http://www.laplace.univ-tlse.fr/
'
' Dépendances :
' Necessite le module "Microsoft Forms 2.0 Object Library" inclus dans la librairie FM20.dll (sous Windows10)
'
' Description :
' Extraction des éléments d'une commande.
' On souhaite récupérer le nom de la personne ayant réaliser la commande,
' le centre de coût (Groupe de recherche), le montant HT, la date et le numéro de la commande.
' =============================================================================

Option Explicit  ' https://docs.microsoft.com/fr-fr/dotnet/visual-basic/language-reference/statements/option-explicit-statement


' Fonction pour l'extraction d'une valeur "strChamp" depuis une chaîne "strIn" en entrée
Function RetourneChamp(strIn As String, strChamp As String) As String
    Dim lngPos As Long ' pour la position du premier caractère de "strIn" dans "strChamp"
    Dim strTemp As String ' chaîne de caractère temporaire pour traitement dans la fonction
    
    lngPos = InStr(strIn, strChamp)
    If lngPos > 0 Then
        strTemp = Mid(strIn, lngPos + Len(strChamp))
        lngPos = InStr(strTemp, vbCr)
        If lngPos > 1 Then
            strTemp = Left(strTemp, lngPos - 1)
            strTemp = Trim(strTemp)
        End If
        strTemp = Trim(strTemp)
    End If
    'MsgBox strTemp & Str(lngPos)  ' pour débogage

    RetourneChamp = strTemp
End Function  ' Fin de la macro RetourneChamp(strIn As String, strChamp As String)


Sub ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)
    Dim i As Long
    
    For i = 1 To Len(strIn)
        If Mid(strIn, i, 1) = car Then Mid(strIn, i, 1) = replaceWith
        Next i
End Sub  ' Fin de la macro ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)


Sub AllerSurLaFeuilleRS()
    Worksheets("RS").Activate
    Range("A1").Activate
End Sub  ' Fin de la macro AllerSurLaFeuilleRS()


Sub AllerSurLaFeuilleFarnell()
    Worksheets("Farnell").Activate
    Range("A1").Activate
End Sub  ' Fin de la macro AllerSurLaFeuilleFarnell()


Sub AllerPremiereLigneVide()
    Dim derniereLigne As Long  ' déclaration de la variable "derniereLigne"
    derniereLigne = Range("A1").End(xlDown).Row  ' récupérer le numéro de la dernière cellule

    If derniereLigne < 1048576 Then
        Range("A" & derniereLigne + 1).Select  ' sélectionne la cellule suivante si "derniereLigne" ne pointe pas la fin du tableau
    Else
        If Range("A1").Value <> "" Then
            Range("A2").Select  ' sélectionne la cellule A2 si la cellule A1 n'est pas vide
        Else
            Range("A1").Select  ' sinon sélectionne la première cellule A1
        End If
    End If
    
End Sub  ' Fin de la macro AllerPremiereLigneVide()


Sub MsgBoxFinMacro()
    Dim Response
    Response = MsgBox("L'extraction automatique du contenu de la commande à échouée !", vbCritical, "Fin de l'Extraction des champs de la commande !")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub AideExtractionInfoFarnell()
    ' Affifhage d'un message d'aide
    MsgBox "- Ouvrez le PDF du mail de <ventes@farnell.com>" & vbCrLf _
    & "   Confirmation de votre commande Farnell - V/Ref :" & vbCrLf & vbCrLf _
    & "- Puis sélectionner tout le texte (CTRL+A) et copier dans le presse papier le contenu avant de cliquer sur le bouton **Commande Farnell** dans Excel.", vbInformation, "Aide du bouton   >>> Commande Farnell <<<"
End Sub  ' Fin de la macro AideExtractionInfoFarnell()


' Extraction des éléments d'une commande
' On doit récupérer le Nom, le centre de coût, le montant HT, la Date, le numéro client
Sub AideExtractionInfoRadioSpares()
    ' Affifhage d'un message d'aide
    ' Dim strOutMsgBox As String
    ' strOutMsgBox = MsgBox("TEXT", vbYesNo + vbQuestion, "TITLE")
    MsgBox "- Copier tout le contenu du mail de <services.clients@rs-components.com> dans le presse papier" & vbCrLf _
    & "   Commande Purchasing Manager - La commande ******* a été saisie sur fr.rs-online.com" & vbCrLf & vbCrLf _
    & "- Puis cliquer sur le bouton **Commande RadioSpares** dans Excel.", vbInformation, "Aide du bouton   >>> Commande RadioSpares <<<"
End Sub  ' Fin de la macro AideExtractionInfoRadioSpares()


Sub ExtractionInfoRadioSpares()
    ' Macro utilisé dans la feuille "=RadioSpares!"
    ' Copier le contenu du mail de <services.clients@rs-components.com> "Commande Purchasing Manager - La commande ******* a été saisie sur fr.rs-online.com"
    ' Puis cliquer sur le bouton **Commande RadioSpares**
    '
    ' Nom : de type chaîne de caractère
    ' Groupe : de type chaîne de caractère
    ' Montant HT : de type numérique (monétaire)
    ' Date : de type date (DD/MM/YYYY)
    ' Numéro Cmd : de type numérique
    If MsgBox("Avez-vous copier le contenu de la commande pour l'extraction automatique ?", vbYesNo + vbQuestion, "Confirmation Extraction d'une commande RadioSpares") = vbYes Then
        Dim MyData As DataObject  ' Permet de récupérer le comptenu du presse-papier (necessite de faire référence à Microsoft Forms 2.0 Object Library)
        Dim strInput As String  ' chaîne en entrée depuis le presse-papier
        'Dim strOutput As String
        Dim strTemp As String  ' chaîne temporaire pour manipulation intermédiaire d'un traitement sur une chaîne
        Dim strNom As String  ' Nom après **Emetteur de besoin**
        Dim strGroupe As String  ' chaîne de caractère du GROUPE
        Dim singleMontantHT As Single  ' nombres décimaux (nombres à virgule)
        Dim dateJJMMAAAA As Date  ' la date de la commande
        Dim longNumCmd As Long  ' le numéro de la commande

        On Error Resume Next
        Set MyData = New DataObject
        MyData.GetFromClipboard
        strInput = MyData.GetText()  ' récupérer la chaîne de caractère du presse-papier (Clipboard)
        Set MyData = Nothing  ' nettoyage de la variable MyData

        AllerPremiereLigneVide

        strNom = RetourneChamp(strInput, "Emetteur de besoin")  ' extraire de "Emetteur de besoinM PRENOM NOM" la chaîne "M PRENOM NOM"
        strNom = RetourneChamp(strNom, " ")  ' rechercher le premier caractère d'espace est supprimer le nbre de caractère du début de la chaîne "M PRENOM NOM", soit la sortie "PRENOM NOM"
        Range("A" & ActiveCell.Row).Value = strNom
    
        strGroupe = RetourneChamp(strInput, "Centre de coût")
        Range("B" & ActiveCell.Row).Value = strGroupe
        
        singleMontantHT = RetourneChamp(strInput, "Montant produits")
        Range("C" & ActiveCell.Row).Value = singleMontantHT

        strTemp = RetourneChamp(strInput, "Date de commande")
        strTemp = Left(strTemp, Len(strTemp) - 6)
        dateJJMMAAAA = CDate(strTemp)
        Range("D" & ActiveCell.Row).Value = dateJJMMAAAA
    
        longNumCmd = RetourneChamp(strInput, "No commande RS")
        Range("E" & ActiveCell.Row).Value = longNumCmd

        MsgBox "Les données de la commande suivante ont été coller" & vbCrLf _
        & "à la dernière ligne du classeur :" & vbCrLf _
        & strNom & vbTab & strGroupe & vbTab & CStr(singleMontantHT) & vbTab & CStr(dateJJMMAAAA) & vbTab & CStr(longNumCmd)
    End If
End Sub  ' Fin de la macro ExtractionInfoRadioSpares()


Sub ExtractionInfoFarnell()
    ' Macro utilisé dans la feuille "=Farnell!"
    ' Ouvrez le PDF du mail de <ventes@farnell.com> "Confirmation de votre commande Farnell - V/Ref :"
    ' Puis faites un copier avant de cliquer sur le bouton **Extraction\nCommande Farnell**
    '
    ' Nom : recherche depuis PDF "Commandé par" ou par le Web "Préparée par:" extraction de "MR PRENOM NOM" de type chaîne de caractère
    ' Groupe : de type chaîne de caractère
    ' Montant HT : de type numérique (monétaire)
    ' Date : de type date (DD/MM/YYYY)
    ' Numéro Cmd : de type numérique
    
    If MsgBox("Avez-vous copier le contenu de la commande pour l'extraction automatique ?", vbYesNo + vbQuestion, "Confirmation Extraction d'une commande Farnell") = vbYes Then
        Dim strInput As String  ' chaîne en entrée depuis le presse-papier
        Dim strOutput As String
        Dim strTemp As String  ' chaîne temporaire pour manipulation intermédiaire d'un traitement sur une chaîne
        
        Dim strNom As String  ' Nom après "Commandé par" ou "Préparée par:"
        Dim strGroupe As String  ' chaîne de caractère du GROUPE
        Dim singleMontantHT As Single  ' nombres décimaux (nombres à virgule)
        Dim dateJJMMAAAA As Date  ' la date de la commande
        Dim strNumCmd As String  ' le numéro de la commande
        Dim strNumCompte As String  ' le numéro de compte

        Dim strA As String, dblremise As Double, dblht As Double
        Dim strB As String
        Dim MyData As DataObject  ' Référence à **Microsoft Forms 2.0 Object Library**
    
        On Error Resume Next  ' Gestion des erreurs, permet de quitter proprement si il y a une erreur
        
        Set MyData = New DataObject
        MyData.GetFromClipboard
        strInput = MyData.GetText()  ' Récupérer la chaîne de caractère du presse-papier (Clipboard)
        
        AllerPremiereLigneVide
        
        ' Recherche du Nom depuis le PDF
        strNom = RetourneChamp(strInput, "Commandé par")  ' Si le champs n'est pas trouvé, la fct retourne une chaîne vide
        If strNom <> "" Then  ' => SI la condition 1 est vraie ALORS
            strNom = RetourneChamp(strNom, "MR")  ' Si strNom est non null rechercher alors "MR"
            Range("A" & ActiveCell.Row).Value = strNom
        Else ' => SINON
            'ElseIf [CONDITION 2] Then ' => SINON, SI la condition 2 est vraie ALORS
            ' Recherche du Nom depuis l'interface Web
            strNom = RetourneChamp(strInput, "Préparée par")  ' Si le champs n'est pas trouvé, la fct retourne une chaîne vide
            If strNom <> "" Then
                Range("A" & ActiveCell.Row).Value = strNom
            Else
                ' Mettre fin à la macro de manière prématurée
                MsgBoxFinMacro
                Exit Sub  ' permet de mettre fin à la macro
            End If
        End If
    
        ' Recherche du GROUPE depuis le PDF
        strGroupe = RetourneChamp(strInput, "GROUPE: ")
        If strGroupe <> "" Then
            Range("B" & ActiveCell.Row).Value = strGroupe
        Else
            ' Recherche du Nom depuis l'interface Web
            strGroupe = RetourneChamp(strInput, "GROUPE")
            If strGroupe <> "" Then
                Range("B" & ActiveCell.Row).Value = strGroupe
            Else
                Exit Sub ' permet de mettre fin à la macro
            End If
        End If
        
        ' Recherche du Montant HT depuis le PDF
        singleMontantHT = RetourneChamp(strInput, "Montant HT")
        Range("C" & ActiveCell.Row).Value = singleMontantHT

        ' Recherche du Montant HT ("Total des marchandises") depuis le Web
        ' TODO...

        ' Recherche de la date depuis le PDF
        strTemp = RetourneChamp(strInput, "Date" & vbCrLf & "1" & vbCrLf)
        dateJJMMAAAA = CDate(strTemp)
        Range("D" & ActiveCell.Row).Value = dateJJMMAAAA

        ' Récupère numéro de commande du PDF
        strTemp = RetourneChamp(strInput, "MR")
        strTemp = Left(strTemp, Len(strTemp) - 0)  ' on cherche le nom
        strTemp = RetourneChamp(strInput, strTemp)  ' on la chaîne après le nom
        strTemp = RetourneChamp(strTemp, vbCr)
        strNumCmd = Replace(strTemp, vbLf, " ")  ' voir si l'espace en plus au début de la chaine strNumCmd n'est pas ajouter ici ?
        Range("E" & ActiveCell.Row).Value = strNumCmd

        ' Récupère le numéro de compte du PDF
        strTemp = RetourneChamp(strInput, "N° de compte")
        strTemp = RetourneChamp(strTemp, vbCr)
        strNumCompte = Replace(strTemp, vbLf, " ")
        Range("F" & ActiveCell.Row).Value = strNumCompte
        
    End If
    
exitCode:
    'MsgBox "Fin de la macro..."
    'clean up code here

End Sub  ' Fin de la macro ExtractionInfoFarnell()

