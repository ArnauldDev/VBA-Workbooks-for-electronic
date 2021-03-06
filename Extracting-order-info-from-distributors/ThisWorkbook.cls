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
' �2021 Arnauld Biganzoli, CNRS-UMR5213-LAPLACE, http://www.laplace.univ-tlse.fr/
'
' D�pendances :
' Necessite le module "Microsoft Forms 2.0 Object Library" inclus dans la librairie FM20.dll (sous Windows10)
'
' Description :
' Extraction des �l�ments d'une commande.
' On souhaite r�cup�rer le nom de la personne ayant r�aliser la commande,
' le centre de co�t (Groupe de recherche), le montant HT, la date et le num�ro de la commande.
' =============================================================================

Option Explicit  ' https://docs.microsoft.com/fr-fr/dotnet/visual-basic/language-reference/statements/option-explicit-statement


' Fonction pour l'extraction d'une valeur "strChamp" depuis une cha�ne "strIn" en entr�e
Function RetourneChamp(strIn As String, strChamp As String) As String
    Dim lngPos As Long ' pour la position du premier caract�re de "strIn" dans "strChamp"
    Dim strTemp As String ' cha�ne de caract�re temporaire pour traitement dans la fonction
    
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
    'MsgBox strTemp & Str(lngPos)  ' pour d�bogage

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
    Dim derniereLigne As Long  ' d�claration de la variable "derniereLigne"
    derniereLigne = Range("A1").End(xlDown).Row  ' r�cup�rer le num�ro de la derni�re cellule

    If derniereLigne < 1048576 Then
        Range("A" & derniereLigne + 1).Select  ' s�lectionne la cellule suivante si "derniereLigne" ne pointe pas la fin du tableau
    Else
        If Range("A1").Value <> "" Then
            Range("A2").Select  ' s�lectionne la cellule A2 si la cellule A1 n'est pas vide
        Else
            Range("A1").Select  ' sinon s�lectionne la premi�re cellule A1
        End If
    End If
    
End Sub  ' Fin de la macro AllerPremiereLigneVide()


Sub MsgBoxFinMacro()
    Dim Response
    Response = MsgBox("L'extraction automatique du contenu de la commande � �chou�e !", vbCritical, "Fin de l'Extraction des champs de la commande !")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub AideExtractionInfoFarnell()
    ' Affifhage d'un message d'aide
    MsgBox "- Ouvrez le PDF du mail de <ventes@farnell.com>" & vbCrLf _
    & "   Confirmation de votre commande Farnell - V/Ref :" & vbCrLf & vbCrLf _
    & "- Puis s�lectionner tout le texte (CTRL+A) et copier dans le presse papier le contenu avant de cliquer sur le bouton **Commande Farnell** dans Excel.", vbInformation, "Aide du bouton   >>> Commande Farnell <<<"
End Sub  ' Fin de la macro AideExtractionInfoFarnell()


' Extraction des �l�ments d'une commande
' On doit r�cup�rer le Nom, le centre de co�t, le montant HT, la Date, le num�ro client
Sub AideExtractionInfoRadioSpares()
    ' Affifhage d'un message d'aide
    ' Dim strOutMsgBox As String
    ' strOutMsgBox = MsgBox("TEXT", vbYesNo + vbQuestion, "TITLE")
    MsgBox "- Copier tout le contenu du mail de <services.clients@rs-components.com> dans le presse papier" & vbCrLf _
    & "   Commande Purchasing Manager - La commande ******* a �t� saisie sur fr.rs-online.com" & vbCrLf & vbCrLf _
    & "- Puis cliquer sur le bouton **Commande RadioSpares** dans Excel.", vbInformation, "Aide du bouton   >>> Commande RadioSpares <<<"
End Sub  ' Fin de la macro AideExtractionInfoRadioSpares()


Sub ExtractionInfoRadioSpares()
    ' Macro utilis� dans la feuille "=RadioSpares!"
    ' Copier le contenu du mail de <services.clients@rs-components.com> "Commande Purchasing Manager - La commande ******* a �t� saisie sur fr.rs-online.com"
    ' Puis cliquer sur le bouton **Commande RadioSpares**
    '
    ' Nom : de type cha�ne de caract�re
    ' Groupe : de type cha�ne de caract�re
    ' Montant HT : de type num�rique (mon�taire)
    ' Date : de type date (DD/MM/YYYY)
    ' Num�ro Cmd : de type num�rique
    If MsgBox("Avez-vous copier le contenu de la commande pour l'extraction automatique ?", vbYesNo + vbQuestion, "Confirmation Extraction d'une commande RadioSpares") = vbYes Then
        Dim MyData As DataObject  ' Permet de r�cup�rer le comptenu du presse-papier (necessite de faire r�f�rence � Microsoft Forms 2.0 Object Library)
        Dim strInput As String  ' cha�ne en entr�e depuis le presse-papier
        'Dim strOutput As String
        Dim strTemp As String  ' cha�ne temporaire pour manipulation interm�diaire d'un traitement sur une cha�ne
        Dim strNom As String  ' Nom apr�s **Emetteur de besoin**
        Dim strGroupe As String  ' cha�ne de caract�re du GROUPE
        Dim singleMontantHT As Single  ' nombres d�cimaux (nombres � virgule)
        Dim dateJJMMAAAA As Date  ' la date de la commande
        Dim longNumCmd As Long  ' le num�ro de la commande

        On Error Resume Next
        Set MyData = New DataObject
        MyData.GetFromClipboard
        strInput = MyData.GetText()  ' r�cup�rer la cha�ne de caract�re du presse-papier (Clipboard)
        Set MyData = Nothing  ' nettoyage de la variable MyData

        AllerPremiereLigneVide

        strNom = RetourneChamp(strInput, "Emetteur de besoin")  ' extraire de "Emetteur de besoinM PRENOM NOM" la cha�ne "M PRENOM NOM"
        strNom = RetourneChamp(strNom, " ")  ' rechercher le premier caract�re d'espace est supprimer le nbre de caract�re du d�but de la cha�ne "M PRENOM NOM", soit la sortie "PRENOM NOM"
        Range("A" & ActiveCell.Row).Value = strNom
    
        strGroupe = RetourneChamp(strInput, "Centre de co�t")
        Range("B" & ActiveCell.Row).Value = strGroupe
        
        singleMontantHT = RetourneChamp(strInput, "Montant produits")
        Range("C" & ActiveCell.Row).Value = singleMontantHT

        strTemp = RetourneChamp(strInput, "Date de commande")
        strTemp = Left(strTemp, Len(strTemp) - 6)
        dateJJMMAAAA = CDate(strTemp)
        Range("D" & ActiveCell.Row).Value = dateJJMMAAAA
    
        longNumCmd = RetourneChamp(strInput, "No commande RS")
        Range("E" & ActiveCell.Row).Value = longNumCmd

        MsgBox "Les donn�es de la commande suivante ont �t� coller" & vbCrLf _
        & "� la derni�re ligne du classeur :" & vbCrLf _
        & strNom & vbTab & strGroupe & vbTab & CStr(singleMontantHT) & vbTab & CStr(dateJJMMAAAA) & vbTab & CStr(longNumCmd)
    End If
End Sub  ' Fin de la macro ExtractionInfoRadioSpares()


Sub ExtractionInfoFarnell()
    ' Macro utilis� dans la feuille "=Farnell!"
    ' Ouvrez le PDF du mail de <ventes@farnell.com> "Confirmation de votre commande Farnell - V/Ref :"
    ' Puis faites un copier avant de cliquer sur le bouton **Extraction\nCommande Farnell**
    '
    ' Nom : recherche depuis PDF "Command� par" ou par le Web "Pr�par�e par:" extraction de "MR PRENOM NOM" de type cha�ne de caract�re
    ' Groupe : de type cha�ne de caract�re
    ' Montant HT : de type num�rique (mon�taire)
    ' Date : de type date (DD/MM/YYYY)
    ' Num�ro Cmd : de type num�rique
    
    If MsgBox("Avez-vous copier le contenu de la commande pour l'extraction automatique ?", vbYesNo + vbQuestion, "Confirmation Extraction d'une commande Farnell") = vbYes Then
        Dim strInput As String  ' cha�ne en entr�e depuis le presse-papier
        Dim strOutput As String
        Dim strTemp As String  ' cha�ne temporaire pour manipulation interm�diaire d'un traitement sur une cha�ne
        
        Dim strNom As String  ' Nom apr�s "Command� par" ou "Pr�par�e par:"
        Dim strGroupe As String  ' cha�ne de caract�re du GROUPE
        Dim singleMontantHT As Single  ' nombres d�cimaux (nombres � virgule)
        Dim dateJJMMAAAA As Date  ' la date de la commande
        Dim strNumCmd As String  ' le num�ro de la commande
        Dim strNumCompte As String  ' le num�ro de compte

        Dim strA As String, dblremise As Double, dblht As Double
        Dim strB As String
        Dim MyData As DataObject  ' R�f�rence � **Microsoft Forms 2.0 Object Library**
    
        On Error Resume Next  ' Gestion des erreurs, permet de quitter proprement si il y a une erreur
        
        Set MyData = New DataObject
        MyData.GetFromClipboard
        strInput = MyData.GetText()  ' R�cup�rer la cha�ne de caract�re du presse-papier (Clipboard)
        
        AllerPremiereLigneVide
        
        ' Recherche du Nom depuis le PDF
        strNom = RetourneChamp(strInput, "Command� par")  ' Si le champs n'est pas trouv�, la fct retourne une cha�ne vide
        If strNom <> "" Then  ' => SI la condition 1 est vraie ALORS
            strNom = RetourneChamp(strNom, "MR")  ' Si strNom est non null rechercher alors "MR"
            Range("A" & ActiveCell.Row).Value = strNom
        Else ' => SINON
            'ElseIf [CONDITION 2] Then ' => SINON, SI la condition 2 est vraie ALORS
            ' Recherche du Nom depuis l'interface Web
            strNom = RetourneChamp(strInput, "Pr�par�e par")  ' Si le champs n'est pas trouv�, la fct retourne une cha�ne vide
            If strNom <> "" Then
                Range("A" & ActiveCell.Row).Value = strNom
            Else
                ' Mettre fin � la macro de mani�re pr�matur�e
                MsgBoxFinMacro
                Exit Sub  ' permet de mettre fin � la macro
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
                Exit Sub ' permet de mettre fin � la macro
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

        ' R�cup�re num�ro de commande du PDF
        strTemp = RetourneChamp(strInput, "MR")
        strTemp = Left(strTemp, Len(strTemp) - 0)  ' on cherche le nom
        strTemp = RetourneChamp(strInput, strTemp)  ' on la cha�ne apr�s le nom
        strTemp = RetourneChamp(strTemp, vbCr)
        strNumCmd = Replace(strTemp, vbLf, " ")  ' voir si l'espace en plus au d�but de la chaine strNumCmd n'est pas ajouter ici ?
        Range("E" & ActiveCell.Row).Value = strNumCmd

        ' R�cup�re le num�ro de compte du PDF
        strTemp = RetourneChamp(strInput, "N� de compte")
        strTemp = RetourneChamp(strTemp, vbCr)
        strNumCompte = Replace(strTemp, vbLf, " ")
        Range("F" & ActiveCell.Row).Value = strNumCompte
        
    End If
    
exitCode:
    'MsgBox "Fin de la macro..."
    'clean up code here

End Sub  ' Fin de la macro ExtractionInfoFarnell()

