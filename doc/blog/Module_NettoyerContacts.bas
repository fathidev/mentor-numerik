Attribute VB_Name = "Module_NettoyerContacts"
' MENTOR NUMERIK - Tuto VBA débutant, épisode 3
' Nettoyer une liste de contacts en 1 clic
' www.mentor-numerik.fr - contact@mentor-numerik.fr
Sub NettoyerContacts()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim i As Long

    ' On cible la feuille Contacts
    Set ws = ThisWorkbook.Worksheets("Contacts")

    ' On cherche la dernière ligne remplie en colonne A
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' On passe sur chaque contact, de la ligne 2 à la dernière
    For i = 2 To derniereLigne

        ' Colonne A : le NOM en majuscules, sans espaces autour
        ws.Cells(i, 1).Value = UCase(Trim(ws.Cells(i, 1).Value))

        ' Colonne B : le Prénom avec une majuscule à chaque début de mot
        ws.Cells(i, 2).Value = Application.WorksheetFunction.Proper(Trim(ws.Cells(i, 2).Value))

        ' Colonne C : l'email en minuscules, sans espaces autour
        ws.Cells(i, 3).Value = LCase(Trim(ws.Cells(i, 3).Value))

    Next i

    MsgBox derniereLigne - 1 & " contacts nettoyés !", vbInformation
End Sub
