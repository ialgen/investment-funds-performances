Attribute VB_Name = "CalculStats"
Option Base 1
Option Explicit

' Fonction Calculant l'equivalent certain des fonds d'investissements
Function EquivalentCertain(r() As Variant, Optional aversion As Double = 3, Optional annualisation As Double = 1) As Double

Dim EC As Double
Dim Er As Double

Er = WorksheetFunction.Average(r())

EC = annualisation * (Er - (aversion / 2) * WorksheetFunction.Var(r()))

EquivalentCertain = EC / 100

End Function

'Fonction calculant le ratio de Sharpe des différents fonds d'investissements
Function Sharpe(r() As Variant, risk_free_rate As Double, Optional annualisation As Double = 1) As Double

'fonction retournant le ratio de Sharpe d'un fonds

'ARGUMENTS:
' - r le vecteur des rendements
' - rf le taux certain

'La formule q calculer
' ratio de Sharpe = racine(annualisation)*(E(r)-rf)/volatilite de r)

Dim Er As Double
Dim std As Double

'calcul des moments (moyenne et ecart-type)
Er = WorksheetFunction.Average(r())
std = WorksheetFunction.StDev(r())

'resultat
Sharpe = Sqr(annualisation) * (Er - risk_free_rate) / std

End Function

'Fonction calculant la volatilite d'un fonds d'investissement
Function volat(r() As Variant, Optional annualisation As Double = 1) As Double

Dim vol As Double

vol = WorksheetFunction.StDev(r()) * Sqr(annualisation)

volat = vol / 100

End Function

'Fonction Calculant le M2
Function MM(r() As Variant, rm() As Variant, risk_free_rate As Double, Optional annualisation As Double = 1) As Variant

'fonction calculant la mesure de performance corrigée du risque M2 (Modigliani & Modigliani)

'ARGUMENTS:
' r le vecteur des rendements du titre (portefeuille)
' rm le vecteur des rendements de l'indice
' rf le taux certain

'M2 = La mesure a calculer = le rendement du portefeuille comprenant le titre et du cash _
    ayant meme volatilité que l'indice de marché
'M2_exc = ecart entre le M2 et le rendement de l'indice (surperformance ˆ volatilite egale ˆ celle du marche)

Dim Er As Double
Dim Erm As Double
Dim volat As Double
Dim volat_m As Double
Dim rapport_volat As Double
Dim M2 As Double
Dim M2_Exc As Double

'calcul des moments (moyennes et ecart-types)
Er = WorksheetFunction.Average(r())
volat = WorksheetFunction.StDev(r())
Erm = WorksheetFunction.Average(rm())
volat_m = WorksheetFunction.StDev(rm())

'calcul du rapport des volatilites
rapport_volat = volat_m / volat

'calcul du M2
M2 = annualisation * (rapport_volat * (Er - risk_free_rate) + risk_free_rate - Erm)

'calcul du M2_exc
M2_Exc = M2 - annualisation * Erm

'resultat
MM = Array(M2 / 100, M2_Exc / 100)

End Function

Function Var(r() As Variant, Optional seuil As Double = 0.05) As Double

'fonction calculant la VAR paramétrique a partir de la série historique r _
des rendements (vecteur colonne) et en fonction du seuil fixé (1%, 5%)

Dim Er As Double
Dim vol As Double

'calcul des moments
Er = WorksheetFunction.Average(r())
vol = WorksheetFunction.StDev(r())

'resultat
Var = WorksheetFunction.Norm_Inv(seuil, Er, vol)

End Function






