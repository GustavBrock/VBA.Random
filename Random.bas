Attribute VB_Name = "Random"
Option Explicit

' Functions for retrieval of truly random numbers.
' 2019-12-26. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.0.
'
' License: MIT.
'
' ReadMe, Disclaimer, and License for data:
'
' Quantum RNG for OpenQu
' ======================
'
' Source and documentation: http://qrng.ethz.ch/
' Hardware: https://www.idquantique.com/random-number-generation/overview/
' Processed by: http://random.openqu.org
'
' The Data is provided "as is" without warranty or any representation of
' accuracy, timeliness or completeness.
'
' Required reference:
'   Microsoft XML, v6.0
' --------------


' Constants.
'
    ' Base URL for the API of "Quantum RNG for OpenQu".
    Private Const UrlApi        As String = "http://random.openqu.org/api/"
    ' Header of Json response if success: {"result":
    Private Const ResultHeader  As String = "{""result"":"

' Enums.
'
    ' Enum for error values for use with Err.Raise.
    Private Enum DtError
        dtInvalidProcedureCallOrArgument = 5
        dtOverflow = 6
        dtTypeMismatch = 13
    End Enum
'

' Retrieves an array of random integer values between a
' minimum and a maximum value.
' By default, only one value of 0 or 1 will be returned.
'
' Arguments:
'   SizeValue:      Count of values retrieved.
'   MinimumValue:   Minimum value that will be retrieved.
'   MaximumValue:   Maximum value that will be retrieved.
'
'   SizeValue should be larger than zero. If not, an array of
'   one element with the value of 0 will be returned.
'   MinimumValue should be smaller than MaximumValue and both
'   should be positive, or unexpected values will be returned.
'
' Acceptable minimum/maximum values are about +/-10E+16.
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnIntegers( _
    Optional SizeValue As Long = 1, _
    Optional MinimumValue As Variant = 0, _
    Optional MaximumValue As Variant = 1) _
    As Variant()
    
    ' Path for returning integer values.
    Const IntegerPath       As String = "randint"
    ' Json response with one value.
    Const NeutralResult     As String = "{""result"": [0]}"
    ' Key names must be lowercase.
    Const SizeKey           As String = "size"
    Const MinimumKey        As String = "min"
    Const MaximumKey        As String = "max"

    Dim Values()            As Variant
    Dim TextValues          As Variant
    Dim MinValue            As Variant
    Dim MaxValue            As Variant
    Dim Index               As Long
    Dim ServiceUrl          As String
    Dim Query               As String
    Dim ResponseText        As String
    Dim Result              As Boolean
    
    If IsNumeric(MinimumValue) And IsNumeric(MaximumValue) Then
        If SizeValue > 0 Then
            ' Round to integer as passing a decimal value will cause the service to fail.
            MinValue = Fix(CDec(MinimumValue))
            MaxValue = Fix(CDec(MaximumValue))
                
            Query = BuildUrlQuery( _
                BuildUrlQueryParameter(SizeKey, SizeValue), _
                BuildUrlQueryParameter(MinimumKey, MinValue), _
                BuildUrlQueryParameter(MaximumKey, MaxValue))
                
            ServiceUrl = UrlApi & IntegerPath & Query
            
            Result = RetrieveDataResponse(ServiceUrl, ResponseText)
        End If
    
        If Result = False Then
            Debug.Print ResponseText
            ResponseText = NeutralResult
        End If
        
        ' Example for ResponseText: {"result": [1, 0, 1]}
        TextValues = Split(Split(Split(ResponseText, "[")(1), "]")(0), ", ")
        ReDim Values(LBound(TextValues) To UBound(TextValues))
        ' Convert the text values to Decimal.
        For Index = LBound(TextValues) To UBound(TextValues)
            Values(Index) = CDec(TextValues(Index))
        Next
    End If
    
    QrnIntegers = Values

End Function

' Retrieves an array of random decimal values that will be
' equal to or larger than 0 (zero) and smaller than 1 (one).
' By default, only one value will be returned.
'
' Arguments:
'   SizeValue:      Count of values retrieved.
'
'   SizeValue should be larger than zero. If not, an array of
'   one element with the value of 0 will be returned.
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnDecimals( _
    Optional SizeValue As Long = 1) _
    As Variant()
    
    ' Path for returning decimal values.
    Const DecimalPath       As String = "rand"
    ' Json response with one value.
    Const NeutralResult     As String = "{""result"": [0.0]}"
    ' Key names must be lowercase.
    Const SizeKey           As String = "size"

    ' Localised decimal separator of a decimal number.
    Dim LocalisedSeparator  As String
    Dim Values()            As Variant
    Dim TextValues          As Variant
    Dim Index               As Long
    Dim ServiceUrl          As String
    Dim Query               As String
    Dim ResponseText        As String
    Dim Result              As Boolean
    
    If SizeValue > 0 Then
        Query = BuildUrlQuery(BuildUrlQueryParameter(SizeKey, SizeValue))
            
        ServiceUrl = UrlApi & DecimalPath & Query
        
        Result = RetrieveDataResponse(ServiceUrl, ResponseText)
    End If
    
    If Result = False Then
        Debug.Print ResponseText
        ResponseText = NeutralResult
    End If
        
    ' Example for ResponseText: {"result": [0.4788413952288314, 0.9344289100110598, 0.6366952682465071]}
    
    ' Find the localised decimal separator.
    LocalisedSeparator = Mid(CStr(0.1), 2, 1)
    
    ' Create array holding the text expressions of the values.
    TextValues = Split(Split(Split(ResponseText, "[")(1), "]")(0), ", ")
    ' Create array to hold the values as Decimal.
    ReDim Values(LBound(TextValues) To UBound(TextValues))
    
    ' Convert the text values to Decimal.
    For Index = LBound(TextValues) To UBound(TextValues)
        ' Replace dot with the localised decimal separator.
        Mid(TextValues(Index), 2) = LocalisedSeparator
        ' Convert the text expression to Decimal.
        Values(Index) = CDec(TextValues(Index))
    Next
    
    QrnDecimals = Values

End Function

' Retrieves one random decimal value that will be equal to or
' larger than 0 (zero) and smaller than 1 (one).
'
' Values will be retrieved from the source in batches to
' relief the burden on the API service and to speed up
' the time to retrieve single values.
'
' The default size of a batch is preset by the constant
' DefaultSize in function QrnDecimalSize.
' The size of the batch (cache) can be preset by calling the function:
'
'   QrnDecimalSize NewCacheSize
'
' Argument Id is for use in a query to force a call of QrnDecimal
' for each record to obtain a random order:
'
'   Select * From SomeTable
'   Order By QrnDecimal([SomeField])
'
' 2019-12-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnDecimal( _
    Optional Id As Variant) _
    As Variant

    Static Values           As Variant
    Static LastIndex        As Long
    
    Dim Value               As Variant
    
    If LastIndex = 0 Then
        ' First run, or all values have been retrieved.
        ' Get size of the cache.
        LastIndex = QrnDecimalSize
        ' Retrieve a new set of values.
        Values = QrnDecimals(LastIndex)
    End If
    
    ' Get the next value.
    ' The index of the array is zero-based.
    LastIndex = LastIndex - 1
    Value = Values(LastIndex)
    
    QrnDecimal = Value

End Function

' Retrieves one random integer value between a minimum and a maximum value.
' By default, a value of 0 or 1 will be returned.
'
' The minimum and maximum values can be preset by calling the functions:
'
'   QrnIntegerMaximum NewMaximumValue
'   QrnIntegerMinimum NewMinimumValue
'
' Acceptable minimum/maximum values are about +/-10E+16.
'
' Example:
'   QrnIntegerMinimum 10
'   QrnIntegerMaximum 20
'   RandomInteger = QrnInteger
'   RandomInteger -> 14
'
' Values will be retrieved from the source in batches to
' relief the burden on the API service and to speed up
' the time to retrieve single values.
'
' The default size of a batch is preset by the constant
' DefaultSize in function QrnIntegerSize.
' The size of the batch (cache) can be preset by calling the function:
'
'   QrnIntegerSize NewCacheSize
'
' Argument Id is for use in a query to force a call of QrnInteger
' for each record to retrieve a random id:
'
'   Select *, QrnInteger([SomeField]) As RandomId
'   From SomeTable
'
' Minimum and/or maximum values of the retrieved ids can be set
' from the query itself, for example 1 and 100 respectively:
'
'   Select *, QrnInteger([SomeField]) As RandomId
'   From SomeTable
'   Where QrnIntegerMinimum(1) > 0 And QrnIntegerMaximum(100) > 0
'
' 2019-12-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnInteger( _
    Optional Id As Variant) _
    As Variant

    Static Values           As Variant
    Static LastIndex        As Long
    Static CurrentMaximum   As Variant
    Static CurrentMinimum   As Variant
    
    Dim Value               As Variant
    
    If CurrentMaximum <> QrnIntegerMaximum Or CurrentMinimum <> QrnIntegerMinimum Then
        ' Reset cache.
        CurrentMaximum = QrnIntegerMaximum
        CurrentMinimum = QrnIntegerMinimum
        LastIndex = 0
    End If
    
    If LastIndex = 0 Then
        ' First run, a reset, or all values have been retrieved.
        ' (Re)set LastIndex to the size of the cache.
        LastIndex = QrnIntegerSize
        ' Retrieve a new set of values.
        Values = QrnIntegers(LastIndex, CurrentMinimum, CurrentMaximum)
    End If
    
    ' Get the next value.
    ' The index of the array is zero-based.
    LastIndex = LastIndex - 1
    Value = Values(LastIndex)
    
    QrnInteger = Value

End Function

' Sets or retrieves the size of the array cached by QrnDecimal.
' To set the size, the new size must be larger than zero.
'
' Example:
'   NewSize = 100
'   QrnDecimalSize NewSize
'   CurrentSize = QrnDecimalSize
'   CurrentSize -> 100
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnDecimalSize( _
    Optional Size As Long) _
    As Long

    Const DefaultSize       As Long = 100
    
    Static CurrentSize      As Long
    
    If Size <= 0 Then
        ' Retrieve cache size.
        If CurrentSize = 0 Then
            ' Cache size has not been set. Use default size.
            CurrentSize = DefaultSize
        End If
    Else
        ' Set cache size.
        CurrentSize = Size
    End If
    
    QrnDecimalSize = CurrentSize

End Function

' Sets or retrieves the size of the array cached by QrnInteger.
' To set the size, the new size must be larger than zero.
'
' Example:
'   NewSize = 100
'   QrnIntegerSize NewSize
'   CurrentSize = QrnIntegerSize
'   CurrentSize -> 100
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnIntegerSize( _
    Optional Size As Long) _
    As Long

    Const DefaultSize       As Long = 100
    
    Static CurrentSize      As Long
    
    If Size <= 0 Then
        ' Retrieve cache size.
        If CurrentSize = 0 Then
            ' Cache size has not been set. Use default size.
            CurrentSize = DefaultSize
        End If
    Else
        ' Set cache size.
        CurrentSize = Size
    End If
    
    QrnIntegerSize = CurrentSize

End Function

' Sets or retrieves the maximum value returned by QrnInteger.
' To set the maximum value, the new value must be larger than zero.
'
' If the current minimum value is larger than the new maximum value,
' the current minimum value will be set to the maximum value - 1.
'
' Example:
'   NewMaximum = 100
'   QrnIntegerMaximum NewMaximum
'   CurrentMaximum = QrnIntegerMaximum
'   CurrentMaximum -> 100
'   ' If the current minimum is 100 or more:
'   CurrentMinimum -> 99
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnIntegerMaximum( _
    Optional MaximumValue As Long) _
    As Long

    Const DefaultMaximum    As Long = 1
    
    Static CurrentMaximum   As Long
    
    If MaximumValue <= 0 Then
        ' Retrieve cache size.
        If CurrentMaximum = 0 Then
            ' Cache size has not been set. Use default maximum.
            CurrentMaximum = DefaultMaximum
        End If
    Else
        ' Set cache size.
        CurrentMaximum = MaximumValue
        ' Avoid a minimum value larger than the maximum value.
        If QrnIntegerMinimum >= CurrentMaximum Then
            QrnIntegerMinimum CurrentMaximum - 1
        End If
    End If
    
    QrnIntegerMaximum = CurrentMaximum

End Function

' Sets or retrieves the minimum value returned by QrnInteger.
' To set the minimum value, the new value must be equal to or larger than zero.
'
' If the current maximum value is lower than the new minimum value,
' the current maximum value will be set to the minimum value + 1.
'
' Example:
'   NewMinimum = 100
'   QrnIntegerMinimum NewMinimum
'   CurrentMinimum = QrnIntegerMinimum
'   CurrentMinimum -> 100
'   ' If the current maximum is 100 or less:
'   CurrentMaximum -> 101
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function QrnIntegerMinimum( _
    Optional MinimumValue As Long = -1) _
    As Long

    Static CurrentMinimum   As Long
    
    If MinimumValue < 0 Then
        ' Retrieve cache size.
    Else
        ' Set cache size.
        CurrentMinimum = MinimumValue
        ' Avoid a maximum values smaller than the minimum value.
        If QrnIntegerMaximum <= CurrentMinimum Then
            QrnIntegerMaximum CurrentMinimum + 1
        End If
    End If
    
    QrnIntegerMinimum = CurrentMinimum

End Function

' Returns a true random number as a Double, like Rnd returns a Single.
' The value will be less than 1 but greater than or equal to zero.
'
' Usage: Excactly like Rnd:
'
'   TrueRandomValue = RndQrn[(Number)]
'
'   Number < 0  ->  The same number every time, using Number as the seed.
'   Number > 0  ->  The next number in the pseudo-random sequence.
'   Number = 0  ->  The most recently generated number.
'   No Number   ->  The next number in the pseudo-random sequence.
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RndQrn( _
    Optional ByVal Number As Single = 1) _
    As Double
    
    Static Value            As Double
    
    Select Case Number
        Case Is > 0 Or (Number = 0 And Value = 0)
            ' Return the next number in the random sequence.
            Value = CDbl(QrnDecimal)
        Case Is = 0
            ' Return the most recently generated number.
        Case Is < 0
            ' Not supported by QRN.
            ' Retrieve value from RndDbl.
            Value = RndDbl(Number)
    End Select
    
    ' Return a value like:
    ' 0.171394365283966
    RndQrn = Value
    
End Function

' Returns a pseudo-random number as a Double, like Rnd returns a Single.
' The value will be less than 1 but greater than or equal to zero.
'
' Usage: Excactly like Rnd:
'
'   PseudoRandomValue = RndDbl[(Number)]
'
'   Number < 0  ->  The same number every time, using Number as the seed.
'   Number > 0  ->  The next number in the pseudo-random sequence.
'   Number = 0  ->  The most recently generated number.
'   No Number   ->  The next number in the pseudo-random sequence.
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RndDbl( _
    Optional ByVal Number As Single = 1) _
    As Double

    ' Exponent to shift the significant digits of a single to
    ' the least significant digits of a double.
    Const Exponent          As Long = 7
    
    Static Value            As Double
    
    Select Case Number
        Case Is <> 0 Or (Number = 0 And Value = 0)
            ' Return the next number in the pseudo-random sequence.
            ' Generate two values like:
            '   0.1851513
            '   0.000000072890967130661
            ' and add these.
            Value = CDbl(Rnd(Number)) + CDbl(Rnd(Number) * 10 ^ -Exponent)
        Case Is = 0
            ' Return the most recently generated number.
    End Select
    
    ' Return a value like:
    '   0.185151372890967
    RndDbl = Value
    
End Function

' Retrieve a Json response from the service URL of the QRN API.
' Retrieved data is returned in parameter ResponseText.
'
' Returns True if success.
'
' Required reference:
'   Microsoft XML, v6.0
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function RetrieveDataResponse( _
    ByVal ServiceUrl As String, _
    ByRef ResponseText As String) _
    As Boolean

    ' ServiceUrl is expected to have URL encoded parameters.
    
    ' Adjustable constants.
    ' Maximum time in seconds to call the service repeatedly
    ' in case of error.
    Const TimeOut           As Integer = 1
    
    ' Fixed constants.
    Const Async             As Boolean = False
    Const StatusOk          As Integer = 200
    Const ErrorNone         As Long = 0
    
    ' Non-caching engine to communicate with the Json service.
    Dim XmlHttp             As New ServerXMLHTTP60
    
    Dim Result              As Boolean
    Dim LastTime            As Date
  
    On Error Resume Next
    
    If ServiceUrl = "" Then
        Err.Raise DtError.dtInvalidProcedureCallOrArgument
    Else
        ' Sometimes a request fails. If so, try a few times more.
        Do
            XmlHttp.Open "GET", ServiceUrl, Async
            XmlHttp.send
            If Err.Number = ErrorNone Then
                Result = True
            Else
                If LastTime = #12:00:00 AM# Then
                    LastTime = Now
                End If
                Debug.Print LastTime, Now
            End If
        Loop Until Result = True Or DateDiff("s", LastTime, Now) > TimeOut
        
        On Error GoTo Err_RetrieveDataResponse
        
        ' Fetch the Json formatted data - or an error message.
        ResponseText = XmlHttp.ResponseText
        
        Select Case XmlHttp.status
            Case StatusOk
                Result = (InStr(ResponseText, ResultHeader) = 1)
            Case Else
                Result = False
        End Select
    End If
    
    RetrieveDataResponse = Result

Exit_RetrieveDataResponse:
    Set XmlHttp = Nothing
    Exit Function

Err_RetrieveDataResponse:
    MsgBox "Error" & Str(Err.Number) & ": " & Err.Description, vbCritical + vbOKOnly, "Web Service Error"
    Resume Exit_RetrieveDataResponse

End Function

' Build the query element of a URL string from a
' parameter array of key/value pairs.
'
' Returns a string of encoded query elements:
'   ?key1=value1&key2=value2& ... &keyN=valueN
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function BuildUrlQuery( _
    ParamArray QueryElements() As Variant) _
    As String
    
    ' Key/Value pairs of QueryElements must be URL encoded.
    
    Const QuerySeparator    As String = "?"
    Const ArgumentSeparator As String = "&"
    
    Dim QueryString         As String
    
    If UBound(QueryElements) > -1 Then
        QueryString = QuerySeparator & Join(QueryElements, ArgumentSeparator)
    End If
    
    BuildUrlQuery = QueryString

End Function

' Build a query element from its key/value pairs.
'
' Returns a key/value string:
'   key=value.
'
' Note:
'   This a reduced function that does not URL encode the key/value pairs.
'
' 2019-12-21. Gustav Brock, Cactus Data ApS, CPH.
'
Private Function BuildUrlQueryParameter( _
    ByVal Key As String, _
    ByVal Value As Variant) _
    As String

    Const KeyValueSeparator As String = "="
    
    Dim QueryElement        As String
    Dim ValueString         As String
    
    ' Trim and URL encode the key/value pair.
    If Trim(Key) <> "" And IsEmpty(Value) = False Then
        ValueString = Trim(CStr(Nz(Value)))
        If ValueString = "" Then
            ValueString = "''"
        End If
        QueryElement = Trim(Key) & KeyValueSeparator & ValueString
    End If
    
    BuildUrlQueryParameter = QueryElement
    
End Function


