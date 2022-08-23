Attribute VB_Name = "RegMulti"
Option Base 1
Option Explicit
Function fnReg(vectY As Variant, vectX As Variant) As Variant

'fonction r�gressant les donn�es de la plage rgY sur celles de la plage rgX et _
reportant les statistiques

fnReg = WorksheetFunction.LinEst(vectY, vectX, True, True)

End Function

Function fnResidu(x As Variant, y As Variant) As Variant

'fonction r�cup�rant le r�sidu de la r�gression de x par rapport � y

Dim r() As Variant, mat()
Dim observ As Integer, i As Integer

observ = UBound(x, 1)
ReDim r(observ, 1)

'r�gression et r�cup�ration du r�sultat
mat = fnReg(x, y)

'calcul du r�sidu
For i = 1 To observ
    r(i, 1) = x(i, 1) - (mat(1, 2) + mat(1, 1) * y(i, 1))
Next i

fnResidu = r

End Function
Function fnOrthog(r As Variant) As Variant

'fonction orthogonalisant les variables successivement les unes par rapport aux autres _
(dans l'ordre des colonnes)

Dim nbre As Integer, observ As Integer
Dim i As Integer, j As Integer, n As Integer
Dim mat() As Variant, x() As Variant, y() As Variant, residu() As Variant

'nombre de variables et d'observations
nbre = UBound(r, 2)
observ = UBound(r, 1)

'redimensionnement des vecteurs et des matrices
ReDim x(observ, 1)
ReDim y(observ, 1)
ReDim mat(observ, nbre)

'r�cup�ration des valeurs initiales des variables
For i = 1 To observ
    For j = 1 To nbre
        mat(i, j) = r(i, j)
    Next j
Next i


'boucle sur les orthonalisations (en nombre=nbre-1)
For n = 1 To nbre - 1
    'r�cup�ration des valeurs de la variable explicative
    For i = 1 To observ
        y(i, 1) = mat(i, n)
    Next i

    For j = n + 1 To nbre
        'r�cup�ration de la variable expliqu�e
        For i = 1 To observ
            x(i, 1) = mat(i, j)
        Next i
        'calcul et r�cup�ration du r�sidu
        residu = fnResidu(x, y)
        'r�cup�ration des valeurs
        For i = 1 To observ
            mat(i, j) = residu(i, 1)
        Next i
    Next j
Next n


fnOrthog = mat

End Function


