Attribute VB_Name = "Module1"
Option Explicit

Public Function RandomNumbers(Upper As Integer, _
   Optional Lower As Integer = 1, _
   Optional HowMany As Integer = 1, _
   Optional Unique As Boolean = True) As Variant
'*******************************************************
    'This Function generates random array of
    'Numbers between Lower & Upper
    'In Addition parameters can include whether
    'UNIQUE values are required
 
   'Note the Result is INCLUSIVE of the Range

    'Debug Example:
    'x = RandomNumbers(49, 1, 7)
    'For n = LBound(x) To UBound(x): Debug.Print x(n);: Next n
    'WARNING HowMany MUST be greater than (Higher - Lower)
    '******************************************************

    On Error GoTo LocalError
    If HowMany > ((Upper + 1) - (Lower - 1)) Then Exit Function
    Dim x           As Integer
    Dim n           As Integer
    Dim arrNums()   As Variant
    Dim colNumbers  As New Collection
    
    ReDim arrNums(HowMany - 1)
    With colNumbers
        'First populate the collection
        For x = Lower To Upper
            .Add x
        Next x
        For x = 0 To HowMany - 1
            n = RandomNumber(0, colNumbers.Count + 1)
            arrNums(x) = colNumbers(n)
            If Unique Then
                colNumbers.Remove n
            End If
        Next x
    End With
    Set colNumbers = Nothing
    RandomNumbers = arrNums
Exit Function
LocalError:
    'Justin (just in case)
    RandomNumbers = ""
End Function


Public Function RandomNumber(Upper As Integer, _
     Lower As Integer) As Integer
    'Generates a Random Number BETWEEN the LOWER and UPPER values
    Randomize
    RandomNumber = Int((Upper - Lower + 1) * Rnd + Lower)
End Function

