Attribute VB_Name = "Performances"
Option Explicit
Option Base 1

Sub perfTable()

Dim wb, wbSource, wbIndices As Workbook
Dim wsPerf, wsSource, wsSource2, wsIndices, wsTampon, wsTamponMulti As Worksheet

Dim wbPath, filename, strat As String

Dim period, nbf, nbfTotal, adress_col As Integer
Dim code, name, celltemp, rangeTampon As Range

Dim currentLine, indiceFond, disp, nbOps, lineTemp, d, facteur As Integer

Dim riskFreeRate As Double
Dim beta As Double
Dim tBeta As Double
Dim R2 As Double
Dim variance As Double

Dim end_date, end_date_rows As Variant
Dim rend(), rendMarket() As Variant
Dim stat(), multi(), res(), rfDate() As Variant
Dim pRisque(), orth() As Variant
Dim vTemp() As Variant

'Initialisation des variables generale
Set wb = ThisWorkbook
wbPath = wb.Path
Set wbIndices = Workbooks.Open(wbPath & "\Datafiles\Indices\indices.xlsb")
Set wsIndices = wbIndices.Worksheets(2)
Set wsPerf = wb.Sheets(1)
wsPerf.name = "Performances"
Set wsTampon = wb.Sheets.Add(after:=wb.Sheets(1))
Set wsTamponMulti = wb.Sheets.Add(after:=wb.Sheets(2))

' CLear le content de la feuille
wsPerf.Cells.Clear

filename = Dir(wbPath & "\Datafiles\*.xlsb", vbNormal)
currentLine = 2

Do Until filename = ""
    Set wbSource = Workbooks.Open(wbPath & "\Datafiles\" & filename)
    Set wsSource = wbSource.Worksheets(1)
    Set wsSource2 = wbSource.Worksheets(2)
  
    period = wsSource2.Cells(Rows.Count, 1).End(xlUp).Row - 1

   'Compteur nombre de fonds d'investissements
    nbf = wsSource.Cells(Rows.Count, 1).End(xlUp).Row - 1

    adress_col = wsSource.UsedRange.Find("Obsolete_Date").Column
    
    indiceFond = 1
    
    ReDim end_date_rows(nbf, 1)
    ReDim end_date(nbf, 1)
    
    For indiceFond = 1 To nbf
        
        'initialisation CAPM
        If IsEmpty(wsSource2.Cells(2, indiceFond + 1)) = False Then
            Set celltemp = wsSource2.Cells(2, indiceFond + 1)
        Else
            Set celltemp = wsSource2.Cells(2, indiceFond + 1).End(xlDown)
            If celltemp.Row > period Then
                GoTo balise
            End If
        End If

        wsPerf.Cells(currentLine + indiceFond - 1, 4).value = wsSource2.Cells(celltemp.Row, 1).value
        riskFreeRate = 0
        
            'remplissage des feuilles tampons
        wsTampon.Cells(1, 1).value = celltemp.value
        wsTampon.Cells(1, 2).value = wsIndices.Cells(celltemp.Row, 4).value
        wsTampon.Cells(1, 3).value = wsIndices.Cells(celltemp.Row, 2).value
        wsTamponMulti.Cells(1, 1).value = celltemp.value
        wsTamponMulti.Cells(1, 2).value = wsIndices.Cells(celltemp.Row, 4).value 'MSCI World
        wsTamponMulti.Cells(1, 3).value = wsIndices.Cells(celltemp.Row, 3).value 'VIX
        wsTamponMulti.Cells(1, 4).value = wsIndices.Cells(celltemp.Row, 6).value - wsIndices.Cells(celltemp.Row, 5).value 'spread growth-value
        wsTamponMulti.Cells(1, 5).value = wsIndices.Cells(celltemp.Row, 10).value - wsIndices.Cells(celltemp.Row, 7).value 'spread risque
        wsTamponMulti.Cells(1, 6).value = wsIndices.Cells(celltemp.Row, 9).value - wsIndices.Cells(celltemp.Row, 8).value 'spread taux
        wsTamponMulti.Cells(1, 7).value = wsIndices.Cells(celltemp.Row, 11).value 'energie
        wsTamponMulti.Cells(1, 8).value = wsIndices.Cells(celltemp.Row, 13).value 'petrole
        wsTamponMulti.Cells(1, 9).value = wsIndices.Cells(celltemp.Row, 14).value 'immobilier
        wsTamponMulti.Cells(1, 10).value = wsIndices.Cells(celltemp.Row, 2).value 'riskfreerate
        riskFreeRate = riskFreeRate + wsIndices.Cells(celltemp.Row, 2).value
        If IsEmpty(wsSource2.Cells(celltemp.Row + 1, indiceFond + 1)) = False Then
            Set celltemp = wsSource2.Cells(celltemp.Row + 1, indiceFond + 1)
        Else
            Set celltemp = celltemp.End(xlDown)
        End If
        
        While celltemp.Row < wsSource2.Rows.Count
            lineTemp = wsTampon.UsedRange.Rows.Count
            wsTampon.Cells(lineTemp + 1, 1).value = celltemp.value
            wsTampon.Cells(lineTemp + 1, 2).value = wsIndices.Cells(celltemp.Row, 4).value
            wsTampon.Cells(lineTemp + 1, 3).value = wsIndices.Cells(celltemp.Row, 2).value
            
            wsTamponMulti.Cells(lineTemp + 1, 1).value = celltemp.value
            wsTamponMulti.Cells(lineTemp + 1, 2).value = wsIndices.Cells(celltemp.Row, 4).value 'MSCI World
            wsTamponMulti.Cells(lineTemp + 1, 3).value = wsIndices.Cells(celltemp.Row, 3).value 'VIX
            wsTamponMulti.Cells(lineTemp + 1, 4).value = wsIndices.Cells(celltemp.Row, 6).value - wsIndices.Cells(celltemp.Row, 5).value 'spread growth-value
            wsTamponMulti.Cells(lineTemp + 1, 5).value = wsIndices.Cells(celltemp.Row, 10).value - wsIndices.Cells(celltemp.Row, 7).value 'spread risque
            wsTamponMulti.Cells(lineTemp + 1, 6).value = wsIndices.Cells(celltemp.Row, 9).value - wsIndices.Cells(celltemp.Row, 8).value 'spread taux
            wsTamponMulti.Cells(lineTemp + 1, 7).value = wsIndices.Cells(celltemp.Row, 11).value 'energie
            wsTamponMulti.Cells(lineTemp + 1, 8).value = wsIndices.Cells(celltemp.Row, 13).value 'petrole
            wsTamponMulti.Cells(lineTemp + 1, 9).value = wsIndices.Cells(celltemp.Row, 14).value 'immobilier
            wsTamponMulti.Cells(lineTemp + 1, 10).value = wsIndices.Cells(celltemp.Row, 2).value 'riskfreerate
            
            riskFreeRate = riskFreeRate + wsIndices.Cells(celltemp.Row, 2).value
            
            If IsEmpty(wsSource2.Cells(celltemp.Row + 1, indiceFond + 1)) = False Then
                Set celltemp = wsSource2.Cells(celltemp.Row + 1, indiceFond + 1)
            Else
                Set celltemp = celltemp.End(xlDown)
            End If
        Wend

        nbOps = wsTampon.Cells(wsTampon.Rows.Count, 1).End(xlUp).Row

            'Calculs de performances
        If nbOps > 12 Then
            riskFreeRate = riskFreeRate / nbOps
            ReDim rend(nbOps)
            rend = wsTampon.Cells(1, 1).Resize(nbOps, 1).value
            rendMarket = wsTampon.Cells(1, 2).Resize(nbOps, 1).value
            
            'calculs CAPM
            stat = WorksheetFunction.LinEst(rend(), rendMarket(), True, True)
            beta = stat(1, 1)
            tBeta = stat(1, 1) / stat(2, 1)
            R2 = stat(3, 1)
            
            rfDate = wsTampon.Cells(1, 3).Resize(nbOps, 1).value
            
            ReDim res(nbOps)
            ReDim pRisque(nbOps)
            
            For d = 1 To nbOps
                res(d) = rend(d, 1) - (rfDate(d, 1) + beta * (rendMarket(d, 1) - rfDate(d, 1)))
                pRisque(d) = rend(d, 1) - rfDate(d, 1)
            Next d
            
            With wsPerf
                .Cells(currentLine + indiceFond - 1, 1).value = wsSource2.Cells(1, indiceFond + 1)
                .Cells(currentLine + indiceFond - 1, 2).value = wsSource.Cells(indiceFond + 1, 2)
                .Cells(currentLine + indiceFond - 1, 7).value = nbOps
                .Cells(currentLine + indiceFond - 1, 8).value = WorksheetFunction.Average(rend()) * 12 / 100
                .Cells(currentLine + indiceFond - 1, 9).value = volat(rend(), 12)
                .Cells(currentLine + indiceFond - 1, 10).value = Sharpe(rend(), riskFreeRate, 12)
                .Cells(currentLine + indiceFond - 1, 11).value = MM(rend(), rendMarket(), riskFreeRate, 12)(1)
                .Cells(currentLine + indiceFond - 1, 12).value = (WorksheetFunction.Average(rend()) * 12 / 100 - 1.5 * volat(rend(), 12) ^ 2)
                .Cells(currentLine + indiceFond - 1, 13).value = WorksheetFunction.Average(pRisque()) * 12 / 100 'Prime de risque
                .Cells(currentLine + indiceFond - 1, 14).value = beta 'Beta
                .Cells(currentLine + indiceFond - 1, 15).value = tBeta 't du Beta
                .Cells(currentLine + indiceFond - 1, 16).value = R2 'R2
                .Cells(currentLine + indiceFond - 1, 17).value = WorksheetFunction.Average(res()) * 12 / 100 'Alpha
                .Cells(currentLine + indiceFond - 1, 18).value = WorksheetFunction.StDev(res()) * Sqr(12) / 100 'risque actif
                .Cells(currentLine + indiceFond - 1, 19).value = (WorksheetFunction.Average(res()) * 12 / 100) / (WorksheetFunction.StDev(res()) * Sqr(12) / 100) 'Ratio d'information
            End With
            
            'calculs regression multifactorielle
                    'orthogonalisation des vecteurs multi
            orth = wsTamponMulti.Cells(1, 2).Resize(nbOps, 8).value
            wsTamponMulti.Cells(1, 2).Resize(nbOps, 8) = fnOrthog(orth())
            orth = wsTamponMulti.Cells(1, 2).Resize(nbOps, 8).value
            
            multi = WorksheetFunction.LinEst(rend(), orth(), True, True)
            ReDim vTemp(8)
            For d = 1 To nbOps
                For facteur = 1 To 8
                    vTemp(facteur) = multi(1, facteur) * wsTamponMulti.Cells(d, 10 - facteur)
                Next facteur
                res(d) = rend(d, 1) - (rfDate(d, 1) + beta * (rendMarket(d, 1) - rfDate(d, 1))) - WorksheetFunction.Sum(vTemp())
            Next d
            
            variance = 0
            For facteur = 1 To 8
                variance = variance + (multi(1, facteur) ^ 2) * WorksheetFunction.Var(wsTamponMulti.Columns(10 - facteur))
            Next facteur
            For facteur = 1 To 8
                vTemp(facteur) = (multi(1, facteur) ^ 2) * WorksheetFunction.Var(wsTamponMulti.Columns(10 - facteur)) / variance
            Next facteur
            
            With wsPerf
                .Cells(currentLine + indiceFond - 1, 20).value = multi(1, 8) 'marche
                .Cells(currentLine + indiceFond - 1, 21).value = multi(1, 7) 'vix
                .Cells(currentLine + indiceFond - 1, 22).value = multi(1, 6) 'spread gw
                .Cells(currentLine + indiceFond - 1, 23).value = multi(1, 5) 'spread credit
                .Cells(currentLine + indiceFond - 1, 24).value = multi(1, 4) 'spread taux
                .Cells(currentLine + indiceFond - 1, 25).value = multi(1, 3) 'energie
                .Cells(currentLine + indiceFond - 1, 26).value = multi(1, 2) 'petrole
                .Cells(currentLine + indiceFond - 1, 27).value = multi(1, 1) 'immobilier
                .Cells(currentLine + indiceFond - 1, 28).value = multi(1, 8) / multi(2, 8)
                .Cells(currentLine + indiceFond - 1, 29).value = multi(1, 7) / multi(2, 7)
                .Cells(currentLine + indiceFond - 1, 30).value = multi(1, 6) / multi(2, 6)
                .Cells(currentLine + indiceFond - 1, 31).value = multi(1, 5) / multi(2, 5)
                .Cells(currentLine + indiceFond - 1, 32).value = multi(1, 4) / multi(2, 4)
                .Cells(currentLine + indiceFond - 1, 33).value = multi(1, 3) / multi(2, 3)
                .Cells(currentLine + indiceFond - 1, 34).value = multi(1, 2) / multi(2, 2)
                .Cells(currentLine + indiceFond - 1, 35).value = multi(1, 1) / multi(2, 1)
                .Cells(currentLine + indiceFond - 1, 36).value = multi(3, 1) 'R2
                .Cells(currentLine + indiceFond - 1, 37).value = WorksheetFunction.Average(res()) * 12 / 100 'Alpha
                .Cells(currentLine + indiceFond - 1, 38).value = WorksheetFunction.StDev(res()) * Sqr(12) / 100 'risque actif
                .Cells(currentLine + indiceFond - 1, 39).value = vTemp(8)
                .Cells(currentLine + indiceFond - 1, 40).value = vTemp(7)
                .Cells(currentLine + indiceFond - 1, 41).value = vTemp(6)
                .Cells(currentLine + indiceFond - 1, 42).value = vTemp(5)
                .Cells(currentLine + indiceFond - 1, 43).value = vTemp(4)
                .Cells(currentLine + indiceFond - 1, 44).value = vTemp(3)
                .Cells(currentLine + indiceFond - 1, 45).value = vTemp(2)
                .Cells(currentLine + indiceFond - 1, 46).value = vTemp(1)
            End With
        End If
        
        'Caractéristiques fonds
        end_date(indiceFond, 1) = wsSource.Cells(indiceFond + 1, adress_col).value
        
        With wsPerf
            .Cells(currentLine + indiceFond - 1, 3).value = Replace(Replace(filename, "hf_", ""), ".xlsb", "")
            If IsEmpty(end_date(indiceFond, 1)) = False Then
                disp = disp + 1
                .Cells(currentLine + indiceFond - 1, 5).value = end_date(indiceFond, 1)
                .Cells(currentLine + indiceFond - 1, 6).value = False
            Else
                .Cells(currentLine + indiceFond - 1, 5).value = "Fond actif"
                .Cells(currentLine + indiceFond - 1, 6).value = True
                end_date_rows(indiceFond, 1) = period + 1
            End If
        
        End With
balise:
        
        wsTampon.UsedRange.Clear
        wsTamponMulti.UsedRange.Clear

    Next indiceFond
    
    currentLine = currentLine + nbf
    Application.DisplayAlerts = False
    wbSource.Close
    Application.DisplayAlerts = True
    filename = Dir()
Loop

Application.DisplayAlerts = False
wsTampon.Delete
wsTamponMulti.Delete
wbIndices.Close
Application.DisplayAlerts = True


'Mise en page
nbfTotal = currentLine - 2

'Ligne des titres
With wsPerf.Cells(1, 1).Resize(1, 46)
    .value = Array("Index des fonds", "Nom des fonds", "Stratégie", "Date de début", "Date de fin", "Survivant", "Nb obs", _
                   "Rendement", "Volatilité", "Ratio de Sharpe", "M2", "EC", _
                   "Prime de risque", "Béta", "t du Beta", "R2 CAPM", "Alpha CAPM", "Risque actif CAPM", "IR", _
                   "sensibilité marché", "sensibilité vix", "sensibilité spread gv", "sensibilité spread credit", "sensibilité spread taux", "sensibilité energie", "sensibilité petrole", "sensibilité immobilier", _
                   "t de sensibilité marché", "t de sensibilité vix", "t de sensibilité spread GV", "t de sensibilité spread credit", "t de sensibilité spread taux", "t de sensibilité energie", "t de sensibilité petrole", "t de sensibilité immobilier", _
                   "R2", "Alpha", "Risque actif", _
                   "part var marché", "part var vix", "part var spread gv", "part var spread credit", "part var spread taux", "part var energie", "part var petrole", "part var immobilier")
    .HorizontalAlignment = xlCenter
    .Borders(xlEdgeBottom).Weight = xlThick
    .Font.name = "Arial"
End With

With wsPerf
    .Cells(2, 8).Resize(nbfTotal, 2).NumberFormat = "0.00%" 'Rendement +vol
    .Cells(2, 10).Resize(nbfTotal, 1).NumberFormat = "0.00" 'Sharpe
    .Cells(2, 11).Resize(nbfTotal, 2).NumberFormat = "0.00%" ' M2 et EC
    .Cells(2, 13).Resize(nbfTotal, 1).NumberFormat = "0.00" 'Prime de risque
    .Cells(2, 14).Resize(nbfTotal, 3).NumberFormat = "0.0" 'Beta et T du beta et R2
    .Cells(2, 17).Resize(nbfTotal, 3).NumberFormat = "0.0%" 'Alpha et risque actif  et Ratio d'information
    .Cells(2, 20).Resize(nbfTotal, 17).NumberFormat = "0.00" 'sensibilites et t des sensibilites
    .Cells(2, 37).Resize(nbfTotal, 2).NumberFormat = "0.00%" 'Alpha et risque actif
    .Cells(2, 39).Resize(nbfTotal, 8).NumberFormat = "0.00%" 'part dans la variance
End With

wsPerf.UsedRange.HorizontalAlignment = xlCenter

'Tri selon le nom des fonds et suppression des dernieres lignes vides
wsPerf.UsedRange.AutoFilter
wsPerf.UsedRange.Sort Key1:=wsPerf.Range("B1"), Order1:=xlAscending, Header:=xlYes

Set celltemp = wsPerf.Cells(wsPerf.Cells(1, 1).End(xlDown).Row + 1, 1).Resize(1000, 100)
celltemp.ClearContents

wsPerf.Columns.AutoFit

End Sub




