Attribute VB_Name = "RandomDemo"
Option Explicit

' Functions for demonstration of truly random numbers.
' 2020-10-17. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.1.
' --------------


' Simulate trows of dice, and return and list the result.
' Calculates and prints the average pip value and its
' offset from the ideal average.
'
' Example:
'   ThrowDice 10, 7
'
'               Die 1         Die 2         Die 3         Die 4         Die 5         Die 6         Die 7         Die 8         Die 9         Die 10
' Throw 1          3             6             3             2             4             1             5             3             3             2
' Throw 2          1             3             1             6             5             1             1             3             2             2
' Throw 3          4             1             1             5             5             3             2             1             4             4
' Throw 4          3             3             6             6             5             3             1             4             6             4
' Throw 5          5             1             6             6             2             6             6             2             4             6
' Throw 6          6             3             1             5             6             4             2             5             6             5
' Throw 7          4             2             5             3             3             1             6             3             2             1
'
' Average pips: 3.50          0,00% off
'
' Note: Even though this example _is_ real, don't expect the average pips to be exactly 3.50.
'
' 2019-12-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function ThrowDice( _
    Optional Throws As Integer = 1, _
    Optional Dice As Integer = 1) _
    As Integer()
    
    ' Array dimensions.
    Const DieDimension      As Long = 1
    Const ThrowDimension    As Long = 2
    
    ' Pip values.
    Const MaximumPip        As Double = 6
    Const MinimumPip        As Double = 1
    ' The average pip equals the median pip.
    Const AveragePip        As Double = (MinimumPip + MaximumPip) / 2
    Const NeutralPip        As Double = 0
    
    Dim DiceTrows()         As Integer
    Dim Die                 As Integer
    Dim Throw               As Integer
    Dim Size                As Long
    Dim Total               As Double
    
    If Dice <= 0 Or Throws <= 0 Then
        ' Return one throw of one die with unknown (neutral) result.
        Throws = 1
        Dice = 1
        Size = 0
    Else
        ' Prepare retrieval of values.
        Size = Throws * Dice
        QrnIntegerSize Size
        QrnIntegerMaximum MaximumPip
        QrnIntegerMinimum MinimumPip
    End If
    
    ReDim DiceTrows(1 To Dice, 1 To Throws)

    If Size > 0 Then
        ' Fill array with results.
        For Throw = LBound(DiceTrows, ThrowDimension) To UBound(DiceTrows, ThrowDimension)
            For Die = LBound(DiceTrows, DieDimension) To UBound(DiceTrows, DieDimension)
                DiceTrows(Die, Throw) = QrnInteger
                Total = Total + DiceTrows(Die, Throw)
            Next
        Next
    End If
    
    ' Print header line.
    Debug.Print , ;
    For Die = LBound(DiceTrows, DieDimension) To UBound(DiceTrows, DieDimension)
        Debug.Print "Die" & Str(Die), ;
    Next
    Debug.Print
    
    ' Print results.
    For Throw = LBound(DiceTrows, ThrowDimension) To UBound(DiceTrows, ThrowDimension)
        Debug.Print "Throw" & Str(Throw);
        For Die = LBound(DiceTrows, DieDimension) To UBound(DiceTrows, DieDimension)
            Debug.Print , "   " & DiceTrows(Die, Throw);
        Next
        Debug.Print
    Next
    Debug.Print
    
    ' Print total.
    If DiceTrows(1, 1) = NeutralPip Then
        ' No total to print.
    Else
        Debug.Print "Average pips:", Format(Total / Size, "0.00"), Format((Total / Size - AveragePip) / AveragePip, "Percent") & " off"
        Debug.Print
    End If
    
    ThrowDice = DiceTrows

End Function

' Quickly sort a Variant array.
'
' The array does not have to be zero- or one-based.
'
' 2018-03-16. Gustav Brock, Cactus Data ApS, CPH.
'
Public Sub QuickSort(ByRef Values As Variant)

    Dim Lows()      As Variant
    Dim Mids()      As Variant
    Dim Tops()      As Variant
    Dim Pivot       As Variant
    Dim Lower       As Long
    Dim Upper       As Long
    Dim UpperLows   As Long
    Dim UpperMids   As Long
    Dim UpperTops   As Long
    
    Dim Value       As Variant
    Dim Item        As Long
    Dim Index       As Long
 
    ' Find count of elements to sort.
    Lower = LBound(Values)
    Upper = UBound(Values)
    If Lower = Upper Then
        ' One element only.
        ' Nothing to do.
        Exit Sub
    End If
    
    
    ' Choose pivot in the middle of the array.
    Pivot = Values(Int((Upper - Lower) / 2) + Lower)
    ' Construct arrays.
    For Each Value In Values
        If Value < Pivot Then
            ReDim Preserve Lows(UpperLows)
            Lows(UpperLows) = Value
            UpperLows = UpperLows + 1
        ElseIf Value > Pivot Then
            ReDim Preserve Tops(UpperTops)
            Tops(UpperTops) = Value
            UpperTops = UpperTops + 1
        Else
            ReDim Preserve Mids(UpperMids)
            Mids(UpperMids) = Value
            UpperMids = UpperMids + 1
        End If
    Next
    
    ' Sort the two split arrays, Lows and Tops.
    If UpperLows > 0 Then
        QuickSort Lows()
    End If
    If UpperTops > 0 Then
        QuickSort Tops()
    End If
    
    ' Concatenate the three arrays and return Values.
    Item = 0
    For Index = 0 To UpperLows - 1
        Values(Lower + Item) = Lows(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperMids - 1
        Values(Lower + Item) = Mids(Index)
        Item = Item + 1
    Next
    For Index = 0 To UpperTops - 1
        Values(Lower + Item) = Tops(Index)
        Item = Item + 1
    Next

End Sub

' Sum the top pip values of a single throw of dice.
' If a top count is not specified, the pip values of all the dice are summed.
'
' 2020-10-17. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function DiceSum( _
    ByVal DieCount As Integer, _
    Optional ByVal TopCount As Integer) _
    As Integer
    
    Dim Throws()    As Integer
    Dim Dice()      As Integer
    Dim TopSum      As Integer
    Dim Index       As Integer
    Dim Count       As Integer
    
    If TopCount <= 0 Then
        TopCount = DieCount
    End If
    
    ' Retrieve two-dimension array with one throw of the dice.
    Throws = ThrowDice(1, DieCount)
    
    ' Convert array to one dimension only
    ReDim Dice(LBound(Throws, 1) To UBound(Throws, 1))
    For Index = LBound(Dice) To UBound(Dice)
        Dice(Index) = Throws(Index, 1)
    Next
    ' Sort the dice by the pip values ascending.
    QuickSort Dice
    
    ' Sum the top pip values
    Index = UBound(Dice)
    While Index >= LBound(Dice) And Count < TopCount
        TopSum = TopSum + Dice(Index)
        Index = Index - 1
        Count = Count + 1
    Wend
    
    DiceSum = TopSum

End Function

