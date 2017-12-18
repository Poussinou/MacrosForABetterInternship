'This file is licensed under the GNU GPL v3 or, at your choice, under the terms of the CECILL-2.1
'Please check https://github.com/Poussinou/MacrosForABetterInternship for updates and licenses

'Option Explicit


Sub Exportation()

Dim LastRowListStage As Long
Dim LastRowListSelection As Long
Dim IDNumber As Long
Dim ToAdd As Boolean

LastRowListStage = Sheets("STAGES 2e année").Range("B" & Rows.Count).End(xlUp).Row
LastRowListSelection = Sheets("Sélection").Range("B" & Rows.Count).End(xlUp).Row

ToAdd = True

For x = 2 To LastRowListStage
    If Sheets("STAGES 2e année").Cells(x, 1).Value = "X" Or Sheets("STAGES 2e année").Cells(x, 1).Value = "x" Then
        Sheets("STAGES 2e année").Cells(x, 1).EntireRow.Copy
        IDNumber = Sheets("STAGES 2e année").Cells(x, 2).Value
        For y = 2 To LastRowListSelection
            If IDNumber = Sheets("Sélection").Cells(y, 2).Value Then
                ToAdd = False
            End If
        Next
        If ToAdd = True Then
            LastRowListSelection = LastRowListSelection + 1
            Sheets("Sélection").Cells(LastRowListSelection, 1).EntireRow.Select
            ActiveSheet.Paste
        End If
    End If
    ToAdd = True
Next

End Sub


Sub Suppression()

Dim LastRowListSelection As Long
LastRowListSelection = Sheets("Sélection").Range("B" & Rows.Count).End(xlUp).Row

For w = 3 To LastRowListSelection
    If Sheets("Sélection").Cells(w, 1).Value <> "x" And Sheets("Sélection").Cells(w, 1).Value <> "X" Then
        Sheets("Sélection").Cells(w, 1).EntireRow.Delete
    End If
Next

End Sub


Sub ApplicationOnGoing()

Dim Row As Long
Row = ActiveCell.Row
Rows(Row).Interior.ColorIndex = 4

End Sub


Sub RefusedApplication()

Dim Row As Long
Row = ActiveCell.Row
Rows(Row).Interior.ColorIndex = 3

End Sub


Sub Help()
    MsgBox "Importer : Cliquez sur ce bouton pour faire apparaitre sur cette page les lignes cochées dans la feuille STAGES 2e année" & Chr(10) & "Nettoyer : Cliquez sur ce bouton pour supprimer les lignes où vous avez enlevé le x dans la colonne des cases à cocher" & Chr(10) & "Candidature en cours : Cliquez sur ce bouton pour mettre la ligne sélectionnée en vert et montrer qu'elle est en cours" & Chr(10) & "Candidature refusée : Cliquez sur ce bouton pour mettre la ligne sélectionnée en rouge et montrer que votre candidature a été refusée"

End Sub
