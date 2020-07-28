Attribute VB_Name = "RandomDemo"
Option Explicit

' Functions for demonstration of truly random numbers.
' 2019-12-26. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.0.
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

