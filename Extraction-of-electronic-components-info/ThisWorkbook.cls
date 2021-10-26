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
' Extraction de caractéristiques de composants.
' On souhaite récupérer les différents paramètres caractérisant un composant
' depuis les sites internet des distributeurs
' =============================================================================

Option Explicit  ' https://docs.microsoft.com/fr-fr/dotnet/visual-basic/language-reference/statements/option-explicit-statement


' Fonction pour l'extraction d'une chaine "strIn" en entrée située avant le paramètre "strChamp"
Function RetourneLigneAvantChamp(strIn As String, strChamp As String) As String
    Dim lngPosDebut As Long ' pour la position du premier caractère de "strIn" dans "strChamp"
    Dim lngPosFin As Long ' pour la position du premier caractère de "strIn" dans "strChamp"
    Dim strTemp As String ' chaîne de caractère temporaire pour traitement dans la fonction
    Dim strTempFin As String ' chaîne de caractère temporaire pour traitement dans la fonction
    
    lngPosFin = InStr(strIn, strChamp) - 2 ' retirer 2 caractères pour aller à la fin de la ligne précédente
    If lngPosFin > 0 Then
        lngPosDebut = InStrRev(strIn, vbCrLf, lngPosFin) ' rechercher l'indice du début de notre chaine de caractère
        strTemp = Mid(strIn, lngPosDebut, lngPosFin - lngPosDebut) ' extraction de la chaine recherché avant strChamp
        strTemp = Replace(strTemp, vbCrLf, "")
    End If

    RetourneLigneAvantChamp = strTemp
End Function  ' Fin de la macro RetourneLigneAvantChamp(strIn As String, strChamp As String) As String


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

    RetourneChamp = strTemp
End Function  ' Fin de la macro RetourneChamp(strIn As String, strChamp As String)


Sub ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)
    Dim i As Long
    
    For i = 1 To Len(strIn)
        If Mid(strIn, i, 1) = car Then Mid(strIn, i, 1) = replaceWith
        Next i
End Sub  ' Fin de la macro ReplaceCar(strIn As String, car As Byte, replaceWith As Byte)


' *****************************************************************************
' Se déplacer dans le classeur

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
    
End Sub  ' fin de la macro AllerPremiereLigneVide()


' *****************************************************************************
' Macro divers
' TODO :
' Test bouton rechercher référence fabricant

Sub MsgBoxFinMacro()
    Dim Response
    Response = MsgBox("L'extraction automatique du contenu de la commande à échouée !", vbCritical, "Fin de l'Extraction des champs de la commande !")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub MsgBoxMacroEnDevenir()
    Dim Response
    Response = MsgBox("Macro en cours de construction...", vbOKOnly, "Exécution d'une Macro")
End Sub  ' Fin de la macro MsgBoxFinMacro()


Sub AccesSiteInternet()
    Dim URL As String
    URL = "https://fr.farnell.com/molex/55932-0810/conn-header-8-voies-1-rangee-2mm/dp/3049278"
    ThisWorkbook.FollowHyperlink URL
End Sub


' *****************************************************************************
' Extraction de caractéristiques de composants.
' Exemple : Description Valeur du composant Unité   Type Boitier    Footprint   Description Fabricant   Référence Fabricant Distributeur    Référence Distributeur  Lien    Catalogue/Feuillet /Lignes  quantité Vrac   quantité Feuillet
'
' Ajout de données supplémentaire pour la gestion du stock du magasin
' - Quantité minimum souhaitée
' - Stock magasin
' - Localisation

Sub ExtractionInfoConnecteurs()
    ' Macro utilisé dans la feuille "=CONNECTEURS!" pour Farnell (Code Commande)
    Dim MyData As DataObject  ' permet de récupérer le comptenu du presse-papier (necessite de faire référence à Microsoft Forms 2.0 Object Library)
    Dim strInput As String  ' chaîne en entrée depuis le presse-papier (attention limité à 255 caractères)
    Dim strTemp As String  ' chaîne temporaire pour manipulation intermédiaire d'un traitement sur une chaîne

    Dim strDescription As String  ' description du composant
    Dim colDescription As String
    colDescription = "A"

    Dim strPas As String  ' chaîne de caractère décrivant le pas du composant en mm (exemple : 2,54)
    Dim colPas As String
    colPas = "B"

    Dim strNombreContacts As String  ' nombre de contacts du composant
    Dim colNombreContacts As String
    colNombreContacts = "C"
    
    Dim strNombreRangees As String  ' nombre de rangées du composant
    Dim colNombreRangees As String
    colNombreRangees = "D"
    
    Dim strOrientationComposant As String  ' Orientation du corps du composant
    Dim colOrientationComposant As String
    colOrientationComposant = "E"

    Dim strTypeMontage As String  ' Type de montage du composant (Traversant ou CMS)
    Dim colTypeMontage As String
    colTypeMontage = "F"

    Dim strTypeRaccordement As String  ' Type de raccordement du composant (à souder ou à sertir)
    Dim colTypeRaccordement As String
    colTypeRaccordement = "G"

    Dim strCourant As String  ' Courant pouvant traverser le composant
    Dim colCourant As String
    colCourant = "H"

    Dim strNomFabricant As String  ' chaîne de caractère du nom du fabricant
    Dim colNomFabricant As String
    colNomFabricant = "I"
    
    Dim strRefFabricant As String  ' chaîne de caractère décrivant la référence du composant
    Dim colRefFabricant As String
    colRefFabricant = "J"
    
    Dim strDistributeur As String
    Dim colDistributeur As String
    colDistributeur = "K"

    Dim strCodeCommande As String  ' chaîne de caractère décrivant le commande du distributeur
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
    strInput = MyData.GetText()  ' récupérer la chaîne de caractère du presse-papier (Clipboard)

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
        ' Recherche pour Farnell : "Réf. Farnell / Réf. Fab. / Mots-clés"
        strDistributeur = RetourneChamp(strInput, "Réf.")
        If strDistributeur = "Farnell / Réf. Fab. / Mots-clés" Then
            Range("A" & ActiveCell.Row).Value = "WIP"
            strDistributeur = "Farnell"
            Range(colDistributeur & ActiveCell.Row).Value = strDistributeur
        Else
            ' Affichage d'un message indicant l'échec de la macro pour l'extraction de données
            MsgBox "Distributeur non trouvé !", vbExclamation, "Recherche pour Connecteur"
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
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
        strNomFabricant = "#N/A"
        Range(colNomFabricant & ActiveCell.Row).Value = strNomFabricant
    End If

    ' Recherche ***RefFabricant*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strRefFabricant = RetourneChamp(strInput, "Référence fabricant:" & vbCrLf)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strRefFabricant = RetourneChamp(strInput, "Imprimer la page" & vbCrLf)
        End If
    End If
    ' Insertion dans la colonne "Référence Fabricant"
    If strRefFabricant <> "" Then
        Range(colRefFabricant & ActiveCell.Row).Value = strRefFabricant
    Else
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
        strCodeCommande = "#N/A"
        Range(colCodeCommande & ActiveCell.Row).Value = strCodeCommande
    End If

    ' Recherche ***Pas*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strInput = RetourneChamp(strInput, "Caractéristiques techniques") ' rechercher uniquement dans la section "Caractéristiques techniques"
        strPas = RetourneChamp(strInput, "Pas") ' rechercher dans la section "Caractéristiques techniques"
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
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
            strNombreContacts = "#N/A"
            Range(colNombreContacts & ActiveCell.Row).Value = strNombreContacts
        End If
    End If
    ' Insertion dans la colonne "Nombre de contacts"
    If Val(strNombreContacts) <> 0 Then
        Range(colNombreContacts & ActiveCell.Row).Value = Val(strNombreContacts)
    Else
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
        strNombreContacts = "#N/A"
        Range(colNombreContacts & ActiveCell.Row).Value = strNombreContacts
    End If
    
    ' Recherche ***NombreRangees*** :
    If strDistributeur = "RS" Then
        ' Recherche pour RS :
        strNombreRangees = RetourneChamp(strInput, "Nombre de rangées") ' Nombre de rangées:
        'Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
    Else ' TODO: recherche if then else if ...
        ' Recherche pour Farnell :
        If strDistributeur = "Farnell" Then
            strNombreRangees = RetourneChamp(strInput, "Nombre de rangées:") ' Nombre de rangées:
            'If strNombreRangees <> "" Then
                'Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
            'Else
                'Range(colNombreRangees & ActiveCell.Row).Value = "#N/A"
            'End If
        'Else
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
            'strNombreRangees = "#N/A"
            'Range(colNombreRangees & ActiveCell.Row).Value = strNombreRangees
        End If
    End If
    ' Insertion dans la colonne "Nombre de rangées"
    If Val(strNombreRangees) <> 0 Then
        Range(colNombreRangees & ActiveCell.Row).Value = Val(strNombreRangees)
    Else
        ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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
            ' Insérer une valeur pour ce champs si l'extraction de données à echouée
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

