Attribute VB_Name = "Exercice"
' ------------------------------------------------------------
' AjouterBoutonTiroir
' Crée un bouton "Appliquer Tiroirs" à l'emplacement prévu s'il n'existe pas
' Référence : appelée manuellement au début
' ------------------------------------------------------------
Sub AjouterBoutonTiroir()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnNom As String
    Dim topPos As Double
    Dim leftPos As Double
    Dim boutonExiste As Boolean

    Set ws = ActiveSheet
    btnNom = "btnTiroir"

    ' Vérifie si le bouton existe déjà (par nom)
    On Error Resume Next
    boutonExiste = Not ws.Shapes(btnNom) Is Nothing
    On Error GoTo 0

    ' Si non présent, on le crée à la cellule J10
    If Not boutonExiste Then
        topPos = ws.Cells(10, 10).Top
        leftPos = ws.Cells(10, 10).Left

        Set btn = ws.Buttons.Add(leftPos, topPos, 150, 30)
        With btn
            .Name = btnNom
            .Caption = "Appliquer Tiroirs"
            .OnAction = "ModifierValeurEtTiroir"
        End With
    End If
End Sub

' ------------------------------------------------------------
' ModifierValeurEtTiroir
' Applique la conversion en kW, les règles des tiroirs et crée le bouton dynamique
' Référence : appelée par le bouton principal
' ------------------------------------------------------------
Sub ModifierValeurEtTiroir()
    Call ModifierValeur
    Call AppliquerTiroir
    Call AjouterBoutonToggleUnite
End Sub

' ------------------------------------------------------------
' ModifierValeur
' Convertit les données de la colonne "Valeur" en kW (division par 1000)
' Renomme la colonne en "Valeur (kW)"
' ------------------------------------------------------------
Sub ModifierValeur()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colValeur As Long
    Dim header As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Recherche de la colonne "Valeur" (pas encore convertie)
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = Trim(LCase(ws.Cells(1, i).Value))
        If header = "valeur" Then
            colValeur = i
            Exit For
        End If
    Next i

    ' Si la colonne est déjà renommée, on considère qu'elle est convertie
    If colValeur = 0 Then
        MsgBox "Les valeurs semblent déjà converties (kW ou W).", vbInformation
        Exit Sub
    End If

    ' Division des données par 1000 pour passer en kW
    For i = 2 To lastRow
        If IsNumeric(ws.Cells(i, colValeur).Value) Then
            ws.Cells(i, colValeur).Value = ws.Cells(i, colValeur).Value / 1000
        End If
    Next i

    ' Mise à jour de l'entête pour signaler la conversion
    ws.Cells(1, colValeur).Value = "Valeur (kW)"
End Sub

' ------------------------------------------------------------
' AppliquerTiroir
' Applique les règles de classification en tiroirs selon jour/heure
' Crée la colonne "Tiroir" si besoin, applique des couleurs et active un filtre
' ------------------------------------------------------------
Sub AppliquerTiroir()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim colValeur As Long, colTiroir As Long
    Dim colDate As Long, colHeure As Long
    Dim dateCell As Date
    Dim heureCell As Variant
    Dim jour As Integer
    Dim h As Integer, m As Integer
    Dim header As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Localisation des colonnes par nom (case-insensitive)
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = Trim(LCase(ws.Cells(1, i).Value))
        Select Case header
            Case "valeur", "valeur (w)", "valeur (kw)": colValeur = i
            Case "date de la mesure": colDate = i
            Case "heure de la mesure": colHeure = i
            Case "tiroir": colTiroir = i
        End Select
    Next i

    ' Vérification de la présence des colonnes essentielles
    If colValeur = 0 Or colDate = 0 Or colHeure = 0 Then
        MsgBox "Colonne 'Valeur', 'Date de la mesure' ou 'Heure de la mesure' introuvable.", vbExclamation
        Exit Sub
    End If

    ' Si la colonne "Tiroir" est absente, on l'insère
    If colTiroir = 0 Then
        ws.Columns(colValeur + 1).Insert Shift:=xlToRight
        ws.Cells(1, colValeur + 1).Value = "Tiroir"
        colTiroir = colValeur + 1
    End If

    ' Attribution des tiroirs selon la logique horaire
    For i = 2 To lastRow
        On Error GoTo ValeurInvalide
        If IsDate(ws.Cells(i, colDate).Value) Then
            dateCell = ws.Cells(i, colDate).Value
            heureCell = TimeValue(ws.Cells(i, colHeure).Text)
            jour = WorksheetFunction.Weekday(dateCell, 2) ' Lundi = 1
            h = Hour(heureCell)
            m = Minute(heureCell)

            If h >= 8 And (h < 20 Or (h = 20 And m = 0)) Then
                If jour >= 1 And jour <= 5 Then
                    ws.Cells(i, colTiroir).Value = "Tiroir 1"
                ElseIf jour = 6 Then
                    ws.Cells(i, colTiroir).Value = "Tiroir 2"
                Else
                    ws.Cells(i, colTiroir).Value = "Tiroir 3"
                End If
            Else
                ws.Cells(i, colTiroir).Value = "Tiroir 3"
            End If
        Else
            ws.Cells(i, colTiroir).Value = "Tiroir 3"
        End If
        GoTo Suite
ValeurInvalide:
        ws.Cells(i, colTiroir).Value = "Tiroir 3"
        Resume Suite
Suite:
        On Error GoTo 0
    Next i

    ' Coloration de la colonne Tiroir uniquement
    For i = 2 To lastRow
        Select Case ws.Cells(i, colTiroir).Value
            Case "Tiroir 1": ws.Cells(i, colTiroir).Interior.Color = RGB(198, 239, 206)
            Case "Tiroir 2": ws.Cells(i, colTiroir).Interior.Color = RGB(255, 235, 156)
            Case "Tiroir 3": ws.Cells(i, colTiroir).Interior.Color = RGB(255, 199, 206)
        End Select
    Next i

    MsgBox "Tiroirs appliqués avec succès !", vbInformation
End Sub

' ------------------------------------------------------------
' AjouterBoutonToggleUnite
' Crée un bouton permettant de basculer entre kW et W
' S'affiche juste en dessous du bouton principal
' ------------------------------------------------------------
Sub AjouterBoutonToggleUnite()
    Dim ws As Worksheet
    Dim btn As Button
    Dim topPos As Double
    Dim leftPos As Double
    Dim btnNom As String

    Set ws = ActiveSheet
    btnNom = "btnToggleUnite"

    ' Supprimer le bouton existant s'il est déjà présent
    On Error Resume Next
    ws.Shapes(btnNom).Delete
    On Error GoTo 0

    ' Position sous le bouton principal "btnTiroir"
    topPos = ws.Shapes("btnTiroir").Top + 35
    leftPos = ws.Shapes("btnTiroir").Left

    ' Création du bouton dynamique
    Set btn = ws.Buttons.Add(leftPos, topPos, 150, 30)
    With btn
        .Name = btnNom
        .Caption = "Repasser en W"
        .OnAction = "ToggleUniteValeur"
    End With
End Sub

' ------------------------------------------------------------
' ToggleUniteValeur
' Alterne entre kW et W dans la colonne "Valeur"
' Modifie le nom de la colonne et le texte du bouton en conséquence
' ------------------------------------------------------------
Sub ToggleUniteValeur()
    Dim ws As Worksheet
    Dim i As Long, colValeur As Long
    Dim header As String
    Dim btn As Shape
    Dim estEnKW As Boolean

    Set ws = ActiveSheet
    Set btn = ws.Shapes("btnToggleUnite")

    ' Recherche de la colonne "Valeur (kW)" ou "Valeur (W)"
    For i = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        header = Trim(LCase(ws.Cells(1, i).Value))
        If header = "valeur (kw)" Or header = "valeur (w)" Then
            colValeur = i
            Exit For
        End If
    Next i

    If colValeur = 0 Then
        MsgBox "Colonne 'Valeur (kW)' ou 'Valeur (W)' introuvable.", vbExclamation
        Exit Sub
    End If

    ' Détecte l'état actuel à partir du nom de colonne
    estEnKW = (Trim(LCase(ws.Cells(1, colValeur).Value)) = "valeur (kw)")

    ' Conversion selon l'état actuel
    For i = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If IsNumeric(ws.Cells(i, colValeur).Value) Then
            If estEnKW Then
                ws.Cells(i, colValeur).Value = ws.Cells(i, colValeur).Value * 1000
            Else
                ws.Cells(i, colValeur).Value = ws.Cells(i, colValeur).Value / 1000
            End If
        End If
    Next i

    ' Mise à jour de l'en-tête et du libellé du bouton
    If estEnKW Then
        ws.Cells(1, colValeur).Value = "Valeur (W)"
        btn.TextFrame.Characters.Text = "Repasser en kW"
    Else
        ws.Cells(1, colValeur).Value = "Valeur (kW)"
        btn.TextFrame.Characters.Text = "Repasser en W"
    End If
End Sub


