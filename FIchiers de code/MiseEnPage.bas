Attribute VB_Name = "MiseEnPage"
Option Explicit
Option Base 1

Sub miseEnPageRecap()

Dim ws2 As Worksheet
Dim cpt As Integer


Set ws2 = ThisWorkbook.Worksheets(2)
''' Mise en forme du tableau !

'' Ligne 1
ws2.Cells(1, 1).Resize(1, 3).value = Array("Stratégie", "Groupe", "eff.")

With ws2.Cells(1, 1).Resize(1, 3)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Bordures des données
With ws2.Cells(3, 2).Resize(18, 37)
    .Borders.Weight = xlThin
End With

''' Mise en place colonne 1
ws2.Cells(3, 1).Resize(3, 1).Merge
With ws2.Cells(3, 1)
    .value = "Toutes stratégies"
    .VerticalAlignment = xlTop
End With

ws2.Cells(6, 1).Resize(3, 1).Merge
With ws2.Cells(6, 1)
    .value = "Stratégies Event-Driven"
    .VerticalAlignment = xlTop
End With

ws2.Cells(9, 1).Resize(3, 1).Merge
With ws2.Cells(9, 1)
    .value = "Stratégies Global Macro"
    .VerticalAlignment = xlTop
End With

ws2.Cells(12, 1).Resize(3, 1).Merge
With ws2.Cells(12, 1)
    .value = "Stratégies Long-Short Equity"
    .VerticalAlignment = xlTop
End With

ws2.Cells(15, 1).Resize(3, 1).Merge
With ws2.Cells(15, 1)
    .value = "Stratégies Merger Arbitrage"
    .VerticalAlignment = xlTop
End With

ws2.Cells(18, 1).Resize(3, 1).Merge
With ws2.Cells(18, 1)
    .value = "Stratégies multi-stratégies"
    .VerticalAlignment = xlTop
End With


''''''''Boucle verticale
For cpt = 0 To 4
    ws2.Cells(2, cpt * 7 + 4).Resize(1, 7).value = Array("Moy", "sd", "5.00%", "25.00%", "50.00%", "75.00%", "95.00%")
    
    With ws2.Cells(2, cpt * 7 + 4).Resize(1, 7)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With ws2.Cells(1, cpt * 7 + 4).Resize(20, 1)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
Next cpt

With ws2.Cells(1, 39).Resize(20, 1)
    .Borders(xlEdgeLeft).Weight = xlThick
End With

''''''''Boucle horizontale
For cpt = 1 To 6
    
    'Colonne B
    ws2.Cells(cpt * 3, 2).Resize(3, 1).value = WorksheetFunction.Transpose(Array("Disparus", "Survivants", "Tous"))
    With ws2.Cells(cpt * 3, 1).Resize(3, 2)
        .HorizontalAlignment = xlLeft
    End With
    
    'Lignes en gras horizontales
    With ws2.Cells(2 + cpt * 3, 1).Resize(1, 38)
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
Next cpt

With ws2.Cells(1, 1).Resize(1, 38)
    .Borders.Weight = xlThick
End With

With ws2.Range("A1:A2")
    .Borders(xlEdgeRight).Weight = xlThin
End With
With ws2.Range("B1:B2")
    .Borders(xlEdgeRight).Weight = xlThin
End With

''''''''' Premiere ligne et stats
''' Espérance de rendement
ws2.Cells(1, 4).Resize(1, 7).Merge
ws2.Cells(1, 4).value = "Espérance de rendement"
With ws2.Cells(1, 4)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Volatilité
ws2.Cells(1, 11).Resize(1, 7).Merge
ws2.Cells(1, 11).value = "Volatilité"
With ws2.Cells(1, 11)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Sharpe ratio
ws2.Cells(1, 18).Resize(1, 7).Merge
ws2.Cells(1, 18).value = "Sharpe Ratio"

With ws2.Cells(1, 18)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' M2
ws2.Cells(1, 25).Resize(1, 7).Merge
ws2.Cells(1, 25).value = "M2"

With ws2.Cells(1, 25)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Equivalent Certain
ws2.Cells(1, 32).Resize(1, 7).Merge
ws2.Cells(1, 32).value = "Equivalent Certain"

With ws2.Cells(1, 32)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With


''' Format des statistiques
With Application.Union(ws2.Cells(3, 4).Resize(18, 14), ws2.Cells(3, 25).Resize(18, 14))
    .NumberFormat = "0.00%"
End With

With ws2.Cells(3, 18).Resize(18, 7)
    .NumberFormat = "0.00"
End With


With ws2.Cells(3, 1).Resize(1, 3)
    .Borders(xlEdgeBottom).Weight = xlThick
End With

Application.DisplayAlerts = False

ws2.Range("1:1").Insert
With ws2.Range("A1:D1")
    .Merge
    .value = "Tableau Récapitulatif de la partie I"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 37
End With

Application.DisplayAlerts = True

''' Autofit des colonnes
ws2.UsedRange.EntireColumn.AutoFit

End Sub

Sub miseEnPageRecapCAPM()

Dim ws2 As Worksheet
Dim cpt As Integer

Set ws2 = ThisWorkbook.Worksheets(3)
''' Mise en forme du tableau !

'' Ligne 1
ws2.Cells(1, 1).Resize(1, 3).value = Array("Stratégie", "Groupe", "eff.")

With ws2.Cells(1, 1).Resize(1, 3)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Bordures des données
With ws2.Cells(3, 2).Resize(18, 44)
    .Borders.Weight = xlThin
End With

''' Mise en place colonne 1
ws2.Cells(3, 1).Resize(3, 1).Merge
With ws2.Cells(3, 1)
    .value = "Toutes stratégies"
    .VerticalAlignment = xlTop
End With

ws2.Cells(6, 1).Resize(3, 1).Merge
With ws2.Cells(6, 1)
    .value = "Stratégies Event-Driven"
    .VerticalAlignment = xlTop
End With

ws2.Cells(9, 1).Resize(3, 1).Merge
With ws2.Cells(9, 1)
    .value = "Stratégies Global Macro"
    .VerticalAlignment = xlTop
End With

ws2.Cells(12, 1).Resize(3, 1).Merge
With ws2.Cells(12, 1)
    .value = "Stratégies Long-Short Equity"
    .VerticalAlignment = xlTop
End With

ws2.Cells(15, 1).Resize(3, 1).Merge
With ws2.Cells(15, 1)
    .value = "Stratégies Merger Arbitrage"
    .VerticalAlignment = xlTop
End With

ws2.Cells(18, 1).Resize(3, 1).Merge
With ws2.Cells(18, 1)
    .value = "Stratégies multi-stratégies"
    .VerticalAlignment = xlTop
End With


''''''''Boucle verticale
For cpt = 0 To 5
    ws2.Cells(2, cpt * 7 + 4).Resize(1, 7).value = Array("Moy", "sd", "5.00%", "25.00%", "50.00%", "75.00%", "95.00%")
    
    With ws2.Cells(2, cpt * 7 + 4).Resize(1, 7)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With ws2.Cells(1, cpt * 7 + 4).Resize(20, 1)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
Next cpt

With ws2.Cells(1, 46).Resize(20, 1)
    .Borders(xlEdgeLeft).Weight = xlThick
End With

''''''''Boucle horizontale
For cpt = 1 To 6
    
    'Colonne B
    ws2.Cells(cpt * 3, 2).Resize(3, 1).value = WorksheetFunction.Transpose(Array("Disparus", "Survivants", "Tous"))
    With ws2.Cells(cpt * 3, 1).Resize(3, 2)
        .HorizontalAlignment = xlLeft
    End With
    
    'Lignes en gras horizontales
    With ws2.Cells(2 + cpt * 3, 1).Resize(1, 45)
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
Next cpt

With ws2.Cells(1, 1).Resize(1, 45)
    .Borders.Weight = xlThick
End With

With ws2.Range("A2")
    .Borders(xlEdgeRight).Weight = xlThin
End With
With ws2.Range("B2")
    .Borders(xlEdgeRight).Weight = xlThin
End With

''''''''' Premiere ligne et stats
''' Prime de risque
ws2.Cells(1, 4).Resize(1, 7).Merge
ws2.Cells(1, 4).value = "Prime de risque"
With ws2.Cells(1, 4)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Béta
ws2.Cells(1, 11).Resize(1, 7).Merge
ws2.Cells(1, 11).value = "Béta"
With ws2.Cells(1, 11)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' t du béta
ws2.Cells(1, 18).Resize(1, 7).Merge
ws2.Cells(1, 18).value = "t du béta"

With ws2.Cells(1, 18)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' R2
ws2.Cells(1, 25).Resize(1, 7).Merge
ws2.Cells(1, 25).value = "R2"

With ws2.Cells(1, 25)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Alpha
ws2.Cells(1, 32).Resize(1, 7).Merge
ws2.Cells(1, 32).value = "Alpha"

With ws2.Cells(1, 32)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Risque actif
ws2.Cells(1, 39).Resize(1, 7).Merge
ws2.Cells(1, 39).value = "Risque actif"

With ws2.Cells(1, 39)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With


''' Format des statistiques
With Application.Union(ws2.Cells(3, 4).Resize(18, 7), ws2.Cells(3, 25).Resize(18, 21))
    .NumberFormat = "0.00%"
End With

With ws2.Cells(3, 11).Resize(18, 14)
    .NumberFormat = "0.00"
End With

Application.DisplayAlerts = False

ws2.Range("1:1").Insert

With ws2.Range("A1:E1")
    .Merge
    .value = "Tableau récapitulatif basé sur le modèle de marché CAPM"
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
    .Interior.ColorIndex = 37
End With

Application.DisplayAlerts = True

''' Autofit des colonnes
ws2.UsedRange.EntireColumn.AutoFit

End Sub

Sub miseEnPageRecapMulti()

Dim ws2 As Worksheet
Dim cpt As Integer
Dim titres As Variant

Set ws2 = ThisWorkbook.Worksheets(4)
''' Mise en forme du tableau !

'' Ligne 1
ws2.Cells(1, 1).Resize(1, 3).value = Array("Stratégie", "Groupe", "eff.")

With ws2.Cells(1, 1).Resize(1, 3)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Bordures des données
With ws2.Cells(3, 2).Resize(18, 58)
    .Borders.Weight = xlThin
End With

''' Mise en place colonne 1
ws2.Cells(3, 1).Resize(3, 1).Merge
With ws2.Cells(3, 1)
    .value = "Toutes stratégies"
    .VerticalAlignment = xlTop
End With

ws2.Cells(6, 1).Resize(3, 1).Merge
With ws2.Cells(6, 1)
    .value = "Stratégies Event-Driven"
    .VerticalAlignment = xlTop
End With

ws2.Cells(9, 1).Resize(3, 1).Merge
With ws2.Cells(9, 1)
    .value = "Stratégies Global Macro"
    .VerticalAlignment = xlTop
End With

ws2.Cells(12, 1).Resize(3, 1).Merge
With ws2.Cells(12, 1)
    .value = "Stratégies Long-Short Equity"
    .VerticalAlignment = xlTop
End With

ws2.Cells(15, 1).Resize(3, 1).Merge
With ws2.Cells(15, 1)
    .value = "Stratégies Merger Arbitrage"
    .VerticalAlignment = xlTop
End With

ws2.Cells(18, 1).Resize(3, 1).Merge
With ws2.Cells(18, 1)
    .value = "Stratégies multi-stratégies"
    .VerticalAlignment = xlTop
End With

''''''''Boucle verticale
For cpt = 0 To 7
    ws2.Cells(2, cpt * 7 + 4).Resize(1, 7).value = Array("Moy", "sd", "5.00%", "25.00%", "50.00%", "75.00%", "95.00%")
    
    With ws2.Cells(2, cpt * 7 + 4).Resize(1, 7)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With ws2.Cells(1, cpt * 7 + 4).Resize(20, 1)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
Next cpt

With ws2.Cells(1, 60).Resize(20, 1)
    .Borders(xlEdgeLeft).Weight = xlThick
End With

''''''''Boucle horizontale
For cpt = 1 To 6
    'Colonne B
    ws2.Cells(cpt * 3, 2).Resize(3, 1).value = WorksheetFunction.Transpose(Array("Disparus", "Survivants", "Tous"))
    With ws2.Cells(cpt * 3, 1).Resize(3, 2)
        .HorizontalAlignment = xlLeft
    End With
    
    'Lignes en gras horizontales
    With ws2.Cells(2 + cpt * 3, 1).Resize(1, 59)
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
Next cpt

With ws2.Cells(1, 1).Resize(1, 59)
    .Borders(xlEdgeBottom).Weight = xlThick
    .Borders(xlEdgeTop).Weight = xlThick
End With

With ws2.Range("A1:A2")
    .Borders(xlEdgeRight).Weight = xlThin
End With
With ws2.Range("B1:B2")
    .Borders(xlEdgeRight).Weight = xlThin
End With

With ws2.Cells(2, 1).Resize(1, 3)
    .Borders(xlEdgeBottom).Weight = xlThick
End With

''''''''' Premiere ligne et stats
titres = Array("Marché", "VIX", "Spread Growth - Value", "Spread Credit", "Spread Taux", "Energie", "Pétrole", "Immobilier")

For cpt = 0 To 7
    ws2.Cells(1, 4 + cpt * 7).Resize(1, 7).Merge
    ws2.Cells(1, 4 + cpt * 7).value = titres(cpt + 1)
    With ws2.Cells(1, cpt * 7 + 4) ''''(cpt + 1) * 7
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
Next cpt

''' Format des statistiques
With ws2.Cells(3, 4).Resize(18, 60)
    .NumberFormat = "0.00"
End With


With ws2.Cells(3, 1).Resize(1, 3)
    .Borders(xlEdgeBottom).Weight = xlThick
End With

''' Autofit des colonnes
ws2.UsedRange.EntireColumn.AutoFit

End Sub

Sub miseEnPageRecapRaMulti()

Dim ws2 As Worksheet
Dim cpt As Integer


Set ws2 = ThisWorkbook.Worksheets(4)
''' Mise en forme du tableau !

'' Ligne 1
ws2.Cells(1, 1).Resize(1, 3).value = Array("Stratégie", "Groupe", "eff.")

With ws2.Cells(1, 1).Resize(1, 3)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Bordures des données
With ws2.Cells(3, 2).Resize(18, 23)
    .Borders.Weight = xlThin
End With

''' Mise en place colonne 1
ws2.Cells(3, 1).Resize(3, 1).Merge
With ws2.Cells(3, 1)
    .value = "Toutes stratégies"
    .VerticalAlignment = xlTop
End With

ws2.Cells(6, 1).Resize(3, 1).Merge
With ws2.Cells(6, 1)
    .value = "Stratégies Event-Driven"
    .VerticalAlignment = xlTop
End With

ws2.Cells(9, 1).Resize(3, 1).Merge
With ws2.Cells(9, 1)
    .value = "Stratégies Global Macro"
    .VerticalAlignment = xlTop
End With

ws2.Cells(12, 1).Resize(3, 1).Merge
With ws2.Cells(12, 1)
    .value = "Stratégies Long-Short Equity"
    .VerticalAlignment = xlTop
End With

ws2.Cells(15, 1).Resize(3, 1).Merge
With ws2.Cells(15, 1)
    .value = "Stratégies Merger Arbitrage"
    .VerticalAlignment = xlTop
End With

ws2.Cells(18, 1).Resize(3, 1).Merge
With ws2.Cells(18, 1)
    .value = "Stratégies multi-stratégies"
    .VerticalAlignment = xlTop
End With


''''''''Boucle verticale
For cpt = 0 To 2
    ws2.Cells(2, cpt * 7 + 4).Resize(1, 7).value = Array("Moy", "sd", "5.00%", "25.00%", "50.00%", "75.00%", "95.00%")
    
    With ws2.Cells(2, cpt * 7 + 4).Resize(1, 7)
        .HorizontalAlignment = xlCenter
        .Borders.Weight = xlMedium
    End With
    
    With ws2.Cells(1, cpt * 7 + 4).Resize(20, 1)
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
Next cpt

With ws2.Cells(1, 24).Resize(20, 1)
    .Borders(xlEdgeRight).Weight = xlThick
End With

''''''''Boucle horizontale
For cpt = 1 To 6
    
    'Colonne B
    ws2.Cells(cpt * 3, 2).Resize(3, 1).value = WorksheetFunction.Transpose(Array("Disparus", "Survivants", "Tous"))
    With ws2.Cells(cpt * 3, 1).Resize(3, 2)
        .HorizontalAlignment = xlLeft
    End With
    
    'Lignes en gras horizontales
    With ws2.Cells(2 + cpt * 3, 1).Resize(1, 24)
        .Borders(xlEdgeBottom).Weight = xlThick
    End With
    
Next cpt

With ws2.Cells(1, 1).Resize(1, 24)
    .Borders.Weight = xlThick
End With

With ws2.Range("A1:A2")
    .Borders(xlEdgeRight).Weight = xlThin
End With
With ws2.Range("B1:B2")
    .Borders(xlEdgeRight).Weight = xlThin
End With

With ws2.Cells(2, 1).Resize(1, 3)
    .Borders(xlEdgeBottom).Weight = xlThick
End With

''''''''' Premiere ligne et stats
''' R2 de la régression
ws2.Cells(1, 4).Resize(1, 7).Merge
ws2.Cells(1, 4).value = "R2 de la régression"
With ws2.Cells(1, 4)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Alpha de Jensen
ws2.Cells(1, 11).Resize(1, 7).Merge
ws2.Cells(1, 11).value = "Alpha de Jensen"
With ws2.Cells(1, 11)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Risque actif
ws2.Cells(1, 18).Resize(1, 7).Merge
ws2.Cells(1, 18).value = "Risque actif"

With ws2.Cells(1, 18)
    .Font.Bold = True
    .HorizontalAlignment = xlCenter
End With

''' Format des statistiques
With ws2.Cells(3, 11).Resize(18, 14)
    .NumberFormat = "0.00%"
End With

With ws2.Cells(3, 4).Resize(18, 7)
    .NumberFormat = "0.00"
End With

''' Autofit des colonnes
ws2.UsedRange.EntireColumn.AutoFit

End Sub
