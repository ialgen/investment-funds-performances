Attribute VB_Name = "Recap"
Option Explicit
Option Base 1

Sub recapTable()

Dim wb As Workbook
Dim wsPerf, wsRecap As Worksheet
Dim rangeTemp As Range
Dim currentLine, nbLine, linePrecS, linePrecE, nbFonds As Integer
Dim cptRecap As Integer

Set wb = ThisWorkbook
Set wsPerf = wb.Worksheets(1)

Set wsRecap = wb.Sheets.Add(after:=wsPerf)
wsRecap.name = "Recap"

wsPerf.UsedRange.Sort Key1:=wsPerf.Range("C1"), Order1:=xlAscending, _
                      key2:=wsPerf.Range("F1"), Order2:=xlAscending, _
                      Header:=xlYes

''''''''Boucle sur tous les fonds permettant d'etablir le tableau pour chaque strategie
nbLine = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row
linePrecS = 2
linePrecE = 2
cptRecap = 6
For currentLine = 3 To nbLine + 1
    
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 8).Resize(nbFonds, 5)

        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 35).value = recapFunction(rangeTemp, 35)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1

        linePrecE = currentLine
    End If

    'Detection de changement de strategie
    If wsPerf.Cells(currentLine, 3).value <> wsPerf.Cells(currentLine - 1, 3).value Then
        nbFonds = currentLine - linePrecS
        Set rangeTemp = wsPerf.Cells(linePrecS, 8).Resize(nbFonds, 5)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 35).value = recapFunction(rangeTemp, 35)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecS = currentLine
    End If
Next currentLine

''''''''Boucle sur tous les fonds pour etablir les resultats englobant toutes les strategies
wsPerf.UsedRange.Sort Key1:=wsPerf.Range("F1"), Order1:=xlAscending, Header:=xlYes
linePrecE = 2
cptRecap = 3
For currentLine = 3 To nbLine + 1
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 8).Resize(nbFonds, 5)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 35).value = recapFunction(rangeTemp, 35)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecE = currentLine
    End If
Next currentLine

nbFonds = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row - 1
Set rangeTemp = wsPerf.Cells(2, 8).Resize(nbFonds, 5)

'gestion du recap
wsRecap.Cells(cptRecap, 4).Resize(1, 35).value = recapFunction(rangeTemp, 35)
wsRecap.Cells(cptRecap, 3).value = nbFonds

Call miseEnPageRecap

End Sub

Sub recapTableCAPM()

Dim wb As Workbook
Dim wsPerf, wsRecap As Worksheet
Dim rangeTemp As Range
Dim currentLine, nbLine, linePrecS, linePrecE, nbFonds As Integer
Dim cptRecap As Integer

Set wb = ThisWorkbook
Set wsPerf = wb.Worksheets(1)

Set wsRecap = wb.Sheets.Add(after:=wb.Worksheets(2))
wsRecap.name = "RecapCAPM"

wsPerf.UsedRange.Sort Key1:=wsPerf.Range("C1"), Order1:=xlAscending, _
                      key2:=wsPerf.Range("F1"), Order2:=xlAscending, _
                      Header:=xlYes

''''''''Boucle sur tous les fonds permettant d'etablir le tableau pour chaque strategie
nbLine = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row
linePrecS = 2
linePrecE = 2
cptRecap = 6
For currentLine = 3 To nbLine + 1
    
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 13).Resize(nbFonds, 6)

        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 42).value = recapFunction(rangeTemp, 42)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1

        linePrecE = currentLine
    End If

    'Detection de changement de strategie
    If wsPerf.Cells(currentLine, 3).value <> wsPerf.Cells(currentLine - 1, 3).value Then
        nbFonds = currentLine - linePrecS
        Set rangeTemp = wsPerf.Cells(linePrecS, 13).Resize(nbFonds, 6)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 42).value = recapFunction(rangeTemp, 42)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecS = currentLine
    End If
Next currentLine

''''''''Boucle sur tous les fonds pour etablir les resultats englobant toutes les strategies
wsPerf.UsedRange.Sort Key1:=wsPerf.Range("F1"), Order1:=xlAscending, Header:=xlYes
linePrecE = 2
cptRecap = 3
For currentLine = 3 To nbLine + 1
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 13).Resize(nbFonds, 6)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 42).value = recapFunction(rangeTemp, 42)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecE = currentLine
    End If
Next currentLine

nbFonds = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row - 1
Set rangeTemp = wsPerf.Cells(2, 13).Resize(nbFonds, 6)

'gestion du recap
wsRecap.Cells(cptRecap, 4).Resize(1, 42).value = recapFunction(rangeTemp, 42)
wsRecap.Cells(cptRecap, 3).value = nbFonds

Call miseEnPageRecapCAPM

End Sub

Sub recapTableMulti(col As Integer)

Dim wb As Workbook
Dim wsPerf, wsRecap As Worksheet
Dim rangeTemp As Range
Dim currentLine, nbLine, linePrecS, linePrecE, nbFonds As Integer
Dim cptRecap As Integer

Set wb = ThisWorkbook
Set wsPerf = wb.Worksheets(1)

If wb.Worksheets.Count < 4 Then
    Set wsRecap = wb.Sheets.Add(after:=wb.Worksheets(3))
    wsRecap.name = "RecapMulti"
Else
    Set wsRecap = wb.Worksheets(4)
End If

wsPerf.UsedRange.Sort Key1:=wsPerf.Range("C1"), Order1:=xlAscending, _
                      key2:=wsPerf.Range("F1"), Order2:=xlAscending, _
                      Header:=xlYes

''''''''Boucle sur tous les fonds permettant d'etablir le tableau pour chaque strategie
nbLine = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row
linePrecS = 2
linePrecE = 2
cptRecap = 6
For currentLine = 3 To nbLine + 1
    
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, col).Resize(nbFonds, 8)

        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 56).value = recapFunction(rangeTemp, 56)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1

        linePrecE = currentLine
    End If

    'Detection de changement de strategie
    If wsPerf.Cells(currentLine, 3).value <> wsPerf.Cells(currentLine - 1, 3).value Then
        nbFonds = currentLine - linePrecS
        Set rangeTemp = wsPerf.Cells(linePrecS, col).Resize(nbFonds, 8)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 56).value = recapFunction(rangeTemp, 56)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecS = currentLine
    End If
Next currentLine

''''''''Boucle sur tous les fonds pour etablir les resultats englobant toutes les strategies
wsPerf.UsedRange.Sort Key1:=wsPerf.Range("F1"), Order1:=xlAscending, Header:=xlYes
linePrecE = 2
cptRecap = 3
For currentLine = 3 To nbLine + 1
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, col).Resize(nbFonds, 8)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 56).value = recapFunction(rangeTemp, 56)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecE = currentLine
    End If
Next currentLine

nbFonds = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row - 1
Set rangeTemp = wsPerf.Cells(2, col).Resize(nbFonds, 8)

'gestion du recap
wsRecap.Cells(cptRecap, 4).Resize(1, 56).value = recapFunction(rangeTemp, 56)
wsRecap.Cells(cptRecap, 3).value = nbFonds

Call miseEnPageRecapMulti

End Sub
Sub recapTableRaMulti(col As Integer)

Dim wb As Workbook
Dim wsPerf, wsRecap As Worksheet
Dim rangeTemp As Range
Dim currentLine, nbLine, linePrecS, linePrecE, nbFonds As Integer
Dim cptRecap As Integer

Set wb = ThisWorkbook
Set wsPerf = wb.Worksheets(1)

If wb.Worksheets.Count < 4 Then
    Set wsRecap = wb.Sheets.Add(after:=wb.Worksheets(3))
    wsRecap.name = "RecapMulti"
Else
    Set wsRecap = wb.Worksheets(4)
End If

wsPerf.UsedRange.Sort Key1:=wsPerf.Range("C1"), Order1:=xlAscending, _
                      key2:=wsPerf.Range("F1"), Order2:=xlAscending, _
                      Header:=xlYes

''''''''Boucle sur tous les fonds permettant d'etablir le tableau pour chaque strategie
nbLine = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row
linePrecS = 2
linePrecE = 2
cptRecap = 6
For currentLine = 3 To nbLine + 1
    
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 36).Resize(nbFonds, 3)

        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 21).value = recapFunction(rangeTemp, 21)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1

        linePrecE = currentLine
    End If

    'Detection de changement de strategie
    If wsPerf.Cells(currentLine, 3).value <> wsPerf.Cells(currentLine - 1, 3).value Then
        nbFonds = currentLine - linePrecS
        Set rangeTemp = wsPerf.Cells(linePrecS, 36).Resize(nbFonds, 3)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 21).value = recapFunction(rangeTemp, 21)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecS = currentLine
    End If
Next currentLine

''''''''Boucle sur tous les fonds pour etablir les resultats englobant toutes les strategies
wsPerf.UsedRange.Sort Key1:=wsPerf.Range("F1"), Order1:=xlAscending, Header:=xlYes
linePrecE = 2
cptRecap = 3
For currentLine = 3 To nbLine + 1
    'Detection de changement d'etat
    If wsPerf.Cells(currentLine, 6).value <> wsPerf.Cells(currentLine - 1, 6).value Then
        nbFonds = currentLine - linePrecE
        Set rangeTemp = wsPerf.Cells(linePrecE, 36).Resize(nbFonds, 3)
        
        'gestion du recap
        wsRecap.Cells(cptRecap, 4).Resize(1, 21).value = recapFunction(rangeTemp, 21)
        wsRecap.Cells(cptRecap, 3).value = nbFonds
        cptRecap = cptRecap + 1
        
        linePrecE = currentLine
    End If
Next currentLine

nbFonds = wsPerf.Cells(wsPerf.Rows.Count, 1).End(xlUp).Row - 1
Set rangeTemp = wsPerf.Cells(2, 36).Resize(nbFonds, 3)

'gestion du recap
wsRecap.Cells(cptRecap, 4).Resize(1, 21).value = recapFunction(rangeTemp, 21)
wsRecap.Cells(cptRecap, 3).value = nbFonds

Call miseEnPageRecapRaMulti

End Sub

Sub recapMulti()

Dim ws As Worksheet

''' Tableau part dans la variance des facteurs de régression
Call recapTableMulti(39)

Set ws = ThisWorkbook.Worksheets(4)
''Titre du tableau
ws.Range("1:1").Insert
With ws.Range("A1:C1")
    .Merge
    .value = "Part dans la variance des paramètres"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 37
End With

''' Tableau T de Student des Sensibilités

ws.Range("1:21").Insert
Call recapTableMulti(28)

''Titre du tableau
ws.Range("1:1").Insert
With ws.Range("A1:C1")
    .Merge
    .value = "T de student des sensibilités"
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
    .Interior.ColorIndex = 37
End With

''' Tableau Sensibilité des paramètres
ws.Range("1:21").Insert
Call recapTableMulti(20)

''Titre du tableau
ws.Range("1:1").Insert
With ws.Range("A1:C1")
    .Merge
    .value = "Sensibilité des paramètres"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 37
End With

''' Tableau Alpha, R2, Risque actif multi-factorielles
ws.Range("1:21").Insert
Call recapTableRaMulti(36)

Application.DisplayAlerts = False

''Titre du tableau
ws.Range("1:1").Insert
With ws.Range("A1:C1")
    .value = "R2, Alpha de Jensen, Risque Actif"
    .Font.Bold = True
    .Merge
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 37
End With

ws.Cells(70, 4).Resize(18, 56).NumberFormat = "0.0%"

ws.name = "Recap Multifactoriel"

Application.DisplayAlerts = True

End Sub

Function recapFunction(rangeTemp As Range, n As Integer) As Variant

Dim result() As Variant
Dim cpt As Integer

ReDim result(n)

For cpt = 0 To n / 7 - 1
    With WorksheetFunction
        result(cpt * 7 + 1) = .Average(rangeTemp.Columns(cpt + 1))
        result(cpt * 7 + 2) = .StDev(rangeTemp.Columns(cpt + 1))
        result(cpt * 7 + 3) = .Percentile(rangeTemp.Columns(cpt + 1), 0.05)
        result(cpt * 7 + 4) = .Percentile(rangeTemp.Columns(cpt + 1), 0.25)
        result(cpt * 7 + 5) = .Percentile(rangeTemp.Columns(cpt + 1), 0.5)
        result(cpt * 7 + 6) = .Percentile(rangeTemp.Columns(cpt + 1), 0.75)
        result(cpt * 7 + 7) = .Percentile(rangeTemp.Columns(cpt + 1), 0.95)
    End With
Next cpt

recapFunction = result()

End Function
