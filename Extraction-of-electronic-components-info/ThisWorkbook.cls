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
' Extraction de caract�ristiques de composants.
' On souhaite r�cup�rer les diff�rents param�tres caract�risant un composant
' depuis les sites internet des distributeurs
' =============================================================================

Option Explicit  ' https://docs.microsoft.com/fr-fr/dotnet/visual-basic/language-reference/statements/option-explicit-statement


' Fonction pour l'extraction d'une chaine "strIn" en entr�e situ�e avant le param�tre "strChamp"
Function RetourneLigneAvantChamp(strIn As String, strChamp As String) As String
    Dim lngPosDebut As Long ' pour la position du premier caract�re de "strIn" dans "strChamp"
    Dim lngPosFin As Long ' pour la position du premier caract�re de "strIn" dans "strChamp"
    Dim strTemp As String ' cha�ne de caract�re temporaire pour traitement dans la fonction
    Dim strTempFin As String ' cha�ne de caract�re temporaire pour traitement dans la fonction
    
    lngPosFin = InStr(strIn, strChamp) - 2 ' retirer 2 caract�res pour aller � la fin de la ligne pr�c�dente
    If lngPosFin > 0 Then
        lngPosDebut = InStrRev(strIn, vbCrLf, lngPosFin) ' rechercher l'indice du d�but de notre chaine de caract�re
        strTemp = Mid(strIn, lngPosDebut, lngPosFin - lngPosDebut) ' extraction de la chaine recherch� avant strChamp
        strTemp = Replace(strTemp, vbCrLf, "")
    End If

    RetourneLigneAvantChamp = strTemp
End Function  ' Fin de la macro RetourneLigneAvantChamp(strIn As String, strChamp As String) As String


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

    RetourneChamp = strTemp
End Function  ' Fin de la macro RetourneChamp(strIn As String, strChamp As String)


Sub ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)
    Dim i As Long
    
    For i = 1 To Len(strIn)
        If Mid(strIn, i, 1) = car Then Mid(strIn, i, 1) = replaceWith
        Next i
End Sub  ' Fin de la macro ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)


' *****************************************************************************
' Se d�placer dans le classeur

Sub AllerSurLaFeuilleCONNECTEURS()
    Worksheets("CONNECTEURS").Activate
    Range("A1").Activate
End Sub  ' Fin de la macro AllerSurLaFeuilleCONNECTEURS()


Sub AllerSurLaFeuilleRESISTANCES()
    Worksheets("RESISTANCES").Activate
    Range("A1").Activate
End Sub  ' Fin de la macro AllerSurLaFeuilleRESISTANCES()


Sub AllerSurLaFeuilleCONDENSATEURS()
    Worksheets("CONDENSATEURS").Activate
    Range("A1").Activate
End Sub  ' Fin de la macro AllerSurLaFeuilleCONDENSATEURS()


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
    
End Sub  ' fin de la macro AllerPremiereLigneVide()


' *****************************************************************************
' Macro divers
' TODO :
' Test bouton rechercher r�f�rence fabricant

Sub MsgBoxFinMacro()
    Dim Response
    Response = MsgBox("L'extraction automatique du contenu de la commande � �chou�e !", vbCritical, "Fin de l'Extraction des champs de la commande !")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub MsgBoxMacroEnDevenir()
    Dim Response
    Response = MsgBox("Macro en cours de construction...", vbOKOnly, "Ex�cution d'une Macro")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub AccesSiteInternet()
    Dim URL As String
    URL = "https://fr.farnell.com/molex/55932-0810/conn-header-8-voies-1-rangee-2mm/dp/3049278"
    ThisWorkbook.FollowHyperlink URL
End Sub


' *****************************************************************************
' Extraction de caract�ristiques de composants.
' Exemple : Description Valeur du composant Unit�   Type Boitier    Footprint   Description Fabricant   R�f�rence Fabricant Distributeur    R�f�rence Distributeur  Lien    Catalogue/Feuillet /Lignes  quantit� Vrac   quantit� Feuillet
'
' Ajout de donn�es suppl�mentaire pour la gestion du stock du magasin
' - Quantit� minimum souhait�e
' - Stock magasin
' - Localisation

Sub ExtractionInfoConnecteurs()
    ' Macro utilis� dans la feuille "=CONNECTEURS!" pour Farnell (Code Commande)
    Dim MyData As DataObject  ' permet de r�cup�rer le comptenu du presse-papier (necessite de faire r�f�rence � Microsoft Forms 2.0 Object Library)
    Dim strInput As String  ' cha�ne en entr�e depuis le presse-papier (attention limit� � 255 caract�res)
    Dim strTemp As String  ' cha�ne temporaire pour manipulation interm�diaire d'un traitement sur une cha�ne

    Dim strDescription As String  ' description du composant
    Dim colDescription As String
    colDescription = "A"

    Dim strPas As String  ' cha�ne de caract�re d�crivant le pas du composant en mm (exemple : 2,54)
    Dim colPas As String
    colPas = "B"

    Dim strNombreContacts As String  ' nombre de contacts du composant
    Dim colNombreContacts As String
    colNombreContacts = "C"
    
    Dim strNombreRangees As String  ' nombre de rang�es du composant
    Dim colNombreRangees As String
    colNombreRangees = "D"
    
    Dim strOrientationComposant As String  ' Orientation du corps du composant
    Dim colOrientationComposant As String
    colOrientationComposant = "E"

    Dim strTypeMontage As String  ' Type de montage du composant (Traversant ou CMS)
    Dim colTypeMontage As String
    colTypeMontage = "F"

    Dim strTypeRaccordement As String  ' Type de raccordement du composant (� souder ou � sertir)
    Dim colTypeRaccordement As String
    colTypeRaccordement = "G"

    Dim strCourant As String  ' Courant pouvant traverser le composant
    Dim colCourant As String
    colCourant = "H"

    Dim strNomFabricant As String  ' cha�ne de caract�re du nom du fabricant
    Dim colNomFabricant As String
    colNomFabricant = "I"
    
    Dim strRefFabricant As String  ' cha�ne de caract�re d�crivant la r�f�rence du composant
    Dim colRefFabricant As String
    colRefFabricant = "J"
    
    Dim strDistributeur As String
    Dim colDistributeur As String
    colDistributeur = "K"

    Dim strCodeCommande As String  ' cha�ne de caract�re d�crivant le commande du distributeur
    Dim colCodeCommande As String
    colCodeCommande = "L"

    Dim colLien As String
    colLien = "M"
    
    Dim colProduitAssocie As String
    colProduitAssocie = "N"
    
    Dim colQteMin As String
    colQteMin = "O"

    Dim colStockMagasin As String
    colStockMagasin = "P"
    
    Dim colLocalisation As String
    colLocalisation = "Q"
 
    On Error Resume Next
    Set MyData = New DataObject
    MyData.GetFromClipboard
    strInput = MyData.GetText()  ' r�cup�rer la cha�ne de caract�re du presse-papier (Clipboard)

    AllerSurLaFeuilleCONNECTEURS
    AllerPremiereLigneVide

    ' Identifiaction du ***Distributeur***
    ' Recherche pour RS : "Code commande RS:"
    strDistributeur = RetourneChamp(strInput, "Code commande")
    If strDistributeur = "RS:" Then
        Range("A" & ActiveCell.Row).Value = "WIP"
        strDistributeur = "RS"
        Range(colDistributeur & ActiveCell.Row).Value = strDistributeur
    Else
        ' Recherche pour Farnell : "R�f. Farnell / R�f. Fab. / Mots-cl�s"
        strDistributeur = RetourneChamp(strInput, "R�f.")
        If strDistributeur = "Farnell / R�f. Fab. / Mots-cl�s" Then
            Range("A" & ActiveCell.Row).Value = "WIP"
            strDistributeur = "Farnell"
            Range(colDistributeur & ActiveCell.Row).Value = strDistributeur
        Else
            ' Affichage d'un message indicant l'�chec de la macro pour l'extraction de donn�es
            MsgBox "Distributeur non trouv� !", vbExclamation, "Recherche pour Connecteur"
            GoTo ProcExit
        End If
    End If
    
    ' Recherche ***NomFabricant*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strNomFabricant = RetourneChamp(strInput, "Marque:" & vbCrLf)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strNomFabricant = RetourneChamp(strInput, "Fabricant :")
        End If
    End If
    ' Insertion dans la colonne "Fabricant"
    If strNomFabricant <> "" Then
        Range(colNomFabricant & ActiveCell.Row).Value = strNomFabricant
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strNomFabricant = "#N/A"
        Range(colNomFabricant & ActiveCell.Row).Value = strNomFabricant
    End If

    ' Recherche ***RefFabricant*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strRefFabricant = RetourneChamp(strInput, "R�f�rence fabricant:" & vbCrLf)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strRefFabricant = RetourneChamp(strInput, "Imprimer la page" & vbCrLf)
        End If
    End If
    ' Insertion dans la colonne "R�f�rence Fabricant"
    If strRefFabricant <> "" Then
        Range(colRefFabricant & ActiveCell.Row).Value = strRefFabricant
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strRefFabricant = "#N/A"
        Range(colRefFabricant & ActiveCell.Row).Value = strRefFabricant
    End If

    ' Recherche ***Description*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strDescription = RetourneLigneAvantChamp(strInput, "Code commande RS")
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strDescription = RetourneChamp(strInput, strRefFabricant & vbCrLf)
        End If
    End If
    ' Insertion dans la colonne "Description"
    If strDescription <> "" Then
        Range(colDescription & ActiveCell.Row).Value = strDescription
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strDescription = "#N/A"
        Range(colDescription & ActiveCell.Row).Value = strDescription
    End If

    ' Recherche ***CodeCommande*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strCodeCommande = RetourneChamp(strInput, "Code commande RS:" & vbCrLf)
    Else
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strCodeCommande = RetourneChamp(strInput, "Code Commande :")
        End If
    End If
    ' Insertion dans la colonne "Code Commande"
    If strCodeCommande <> "" Then
        Range(colCodeCommande & ActiveCell.Row).Value = strCodeCommande
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strCodeCommande = "#N/A"
        Range(colCodeCommande & ActiveCell.Row).Value = strCodeCommande
    End If

    ' Recherche ***Pas*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strInput = RetourneChamp(strInput, "Caract�ristiques techniques") ' rechercher uniquement dans la section "Caract�ristiques techniques"
        strPas = RetourneChamp(strInput, "Pas") ' rechercher dans la section "Caract�ristiques techniques"
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strPas = RetourneChamp(strInput, "Pas:")
            Range(colPas & ActiveCell.Row).Value = Val(strPas)
        End If
    End If
    ' Insertion dans la colonne "Pas (mm)"
    If Val(strPas) <> 0 Then
        Range(colPas & ActiveCell.Row).Value = Val(strPas)
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strPas = "#N/A"
        Range(colPas & ActiveCell.Row).Value = strPas
    End If
        
    ' Recherche ***NombreContacts*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strNombreContacts = RetourneChamp(strInput, "Nombre de contacts")
        'Range(colNombreContacts & ActiveCell.Row).Value = Val(strNombreContacts)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strNombreContacts = RetourneChamp(strInput, "Nombre de contacts:")
            If strNombreContacts <> "" Then
            Else
                strNombreContacts = RetourneChamp(strInput, "Nbre de contacts:")
                If strNombreContacts <> "" Then
                Else
                    strNombreContacts = RetourneChamp(strInput, "Nombre de Positions:")
                End If
            End If
        Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            strNombreContacts = "#N/A"
            Range(colNombreContacts & ActiveCell.Row).Value = strNombreContacts
        End If
    End If
    ' Insertion dans la colonne "Nombre de contacts"
    If Val(strNombreContacts) <> 0 Then
        Range(colNombreContacts & ActiveCell.Row).Value = Val(strNombreContacts)
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strNombreContacts = "#N/A"
        Range(colNombreContacts & ActiveCell.Row).Value = strNombreContacts
    End If
    
    ' Recherche ***NombreRangees*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strNombreRangees = RetourneChamp(strInput, "Nombre de rang�es") ' Nombre de rang�es:
        'Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strNombreRangees = RetourneChamp(strInput, "Nombre de rang�es:") ' Nombre de rang�es:
            'If strNombreRangees <> "" Then
                'Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
            'Else
                'Range(colNombreRangees & ActiveCell.Row).Value = "#N/A"
            'End If
        'Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            'strNombreRangees = "#N/A"
            'Range(colNombreRangees & ActiveCell.Row).Value = strNombreRangees
        End If
    End If
    ' Insertion dans la colonne "Nombre de rang�es"
    If Val(strNombreRangees) <> 0 Then
        Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
    Else
        ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
        strNombreRangees = "#N/A"
        Range(colNombreRangees & ActiveCell.Row).Value = strNombreRangees
    End If

    ' Recherche ***OrientationComposant*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strOrientationComposant = RetourneChamp(strInput, "Orientation du corps")
        Range(colOrientationComposant & ActiveCell.Row).Value = strOrientationComposant
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strOrientationComposant = RetourneChamp(strInput, "Orientation du corps du connecteur:")
            If strOrientationComposant <> "" Then
                Range(colOrientationComposant & ActiveCell.Row).Value = strOrientationComposant
            Else
                strOrientationComposant = RetourneChamp(strInput, "Type de terminaison de contact:")
                If strOrientationComposant <> "" Then
                    Range(colOrientationComposant & ActiveCell.Row).Value = strOrientationComposant
                Else
                    Range(colOrientationComposant & ActiveCell.Row).Value = "#N/A"
                End If
            End If
        Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            strOrientationComposant = "#N/A"
            Range(colOrientationComposant & ActiveCell.Row).Value = strOrientationComposant
        End If
    End If

    ' Recherche ***TypeMontage*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strTypeMontage = RetourneChamp(strInput, "Type de montage")
        Range(colTypeMontage & ActiveCell.Row).Value = strTypeMontage
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strTypeMontage = RetourneChamp(strInput, "Montage connecteur:")
            Range(colTypeMontage & ActiveCell.Row).Value = strTypeMontage
        Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            strTypeMontage = "#N/A"
            Range(colTypeMontage & ActiveCell.Row).Value = strTypeMontage
        End If
    End If

    ' Recherche ***TypeRaccordement*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strTypeRaccordement = RetourneChamp(strInput, "Type de raccordement")
        Range(colTypeRaccordement & ActiveCell.Row).Value = strTypeRaccordement
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        'If strDistributeur = "Farnell" Then

        'Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            'strTypeRaccordement = "TODO"
            'Range(colTypeRaccordement & ActiveCell.Row).Value = strTypeRaccordement
        'End If
    End If
        
    ' Recherche ***Courant*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strCourant = RetourneChamp(strInput, "Courant")
        Range(colCourant & ActiveCell.Row).Value = Val(strCourant)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            'strCourant = RetourneChamp(strInput, "")
            'Range(colCourant & ActiveCell.Row).Value = strCourant
        Else
            ' Ins�rer une valeur pour ce champs si l'extraction de donn�es � echou�e
            strCourant = "#N/A"
            Range(colCourant & ActiveCell.Row).Value = strCourant
        End If
    End If

    Range(colLien & ActiveCell.Row).Value = "TODO"
    Range(colProduitAssocie & ActiveCell.Row).Value = "#N/A"
    Range(colQteMin & ActiveCell.Row).Value = "TODO"
    Range(colStockMagasin & ActiveCell.Row).Value = "TODO"
    Range(colLocalisation & ActiveCell.Row).Value = "TODO"
        
ProcExit:
    Set MyData = Nothing 'nettoyage de la variable MyData
End Sub  ' Fin de la macro ExtractionInfoConnecteurs()


Sub ExtractionInfoResistances()
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

