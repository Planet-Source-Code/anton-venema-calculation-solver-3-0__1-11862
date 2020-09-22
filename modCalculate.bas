Attribute VB_Name = "modCalculate"
Option Explicit

Private Char As String
Private CurrentEntryIndex As Integer
Private ErrorMessage As String
Private Help As Boolean
Private InError As Boolean
Private InputString As String
Private OutputString As String
Private OutputValue As Double
Private Value As Double
Private ValueString As String

Public DecIndex As Integer
Public PrevAnswer As Double
Public PrevEntry As String

Const Pi = 3.14159265358979
''''''''''''''''''''''''''''''''''''''
'      Mathematical Grammar Key      '
'          (Hierarchy Chart)         '
''''''''''''''''''''''''''''''''''''''
'                                    '
'  E ::=  T | T + E | T - E          '
'  T ::=  F | F * T | F / T          '
'  F ::=  Number | ( E )             '
'                                    '
''''''''''''''''''''''''''''''''''''''
'                                    '
'  Final Result:            E        '
'                           |        '
'                           T        '
'                          /|\       '
'                         F \ T      '
'                        /|\ \ \     '
'                       / E \ \ \    '
'                      / /|\ \ \ \   '
'                     / T | E \ \ F  '
'                     | | | | | | |  '
'                     | | | T | | |  '
'                     | | | | | | |  '
'                     | F | F | | |  '
'                     | | | | | | |  '
'  Base Equation:     | 1 + 2 ) * 3  '
'                                    '
''''''''''''''''''''''''''''''''''''''
'                                    '
'  Example:                          '
'                                    '
'  Final Result:            9        '
'                           |        '
'                           9        '
'                          /|\       '
'                         3 * 3      '
'                        /|\ \ \     '
'                       ( 3 ) * 3    '
'                      / /|\ \ \ \   '
'                     ( 1 + 2 ) * 3  '
'                     | | | | | | |  '
'                     ( 1 + 2 ) * 3  '
'                     | | | | | | |  '
'  Base Equation:     ( 1 + 2 ) * 3  '
'                                    '
''''''''''''''''''''''''''''''''''''''

Public Sub CalculateEntry()
'On Error GoTo ErrorHandler:
Dim Answer As String
Dim BinAnswer As String
Dim DecimalCheck As Long
Dim i As Integer
Dim LenAfterDecimal As Long
Dim NumOfDecimals As Integer
Dim Remainder As String
Dim Tag As String

    'Set default values
    CurrentEntryIndex = 1
    Help = False
    InError = False
    InputString = frmCalcSolver.txtEntry.Text
    PrevEntry = frmCalcSolver.txtEntry.Text

    'Extract the first token
    ExtractToken

    'Evaluate the entire expression
    Answer = CStr(GetE)

    'Open Syntax help
    If Help Then
        frmCalcSolver.mnuHelpHelp_Click
        Exit Sub
    End If

    'If we "finished" the evaluation prematurely, an
    'error occured
    If Not InError And OutputString <> "EOS" Then
        TrapErrors 0
    End If

    'Set error message if error occurred
    If InError Then
        Answer = ">> " + ErrorMessage + vbNewLine + frmCalcSolver.txtAnswer.Text

    Else

        'Set previous answer
        PrevAnswer = Answer
        Tag = ""
        If frmCalcSolver.optBaseMode(1).Value = True Then

            'Convert to binary if necessary
            If CDbl(Answer) <= 32767 Then
                BinAnswer = ""
                DecimalCheck = InStr(1, CStr(Answer), ".")
                If DecimalCheck <> 0 Then
                    If CInt(Mid(CStr(Answer), DecimalCheck + 1, 1)) < 5 Then
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                    Else
                        Answer = CDbl(Left(Answer, DecimalCheck - 1)) + 1
                    End If
                End If
                Do
                    Answer = Answer / 2
                    DecimalCheck = InStr(1, CStr(Answer), ".")
                    If DecimalCheck = 0 Then
                        Remainder = "0"
                    Else
                        Answer = CDbl(Left(Answer, DecimalCheck - 1))
                        Remainder = "1"
                    End If
                    BinAnswer = Remainder + BinAnswer
                Loop Until Answer < 1
                Answer = CDbl(BinAnswer)
                Tag = " (bin)"
            End If
        ElseIf frmCalcSolver.optBaseMode(2).Value = True Then

            'Convert to hexadecimal if necessary
            Answer = Hex(Answer)
            Tag = " (hex)"
        ElseIf frmCalcSolver.optBaseMode(3).Value = True Then

            'Convert to octadecimal if necessary
            Answer = Oct(Answer)
            Tag = " (oct)"
        Else

            'If in decimal mode, convert to set
            'number of decimal places
            If frmCalcSolver.txtDecimal.Text <> "F" Then

                'Check for decimal
                NumOfDecimals = Val(frmCalcSolver.txtDecimal.Text)
                DecimalCheck = InStr(1, CStr(Answer), ".")

                'If decimal does not exist, tag on the number
                'of zeroes that the user specified
                If DecimalCheck = 0 Then
                    If NumOfDecimals <> "0" Then
                        Answer = Answer + "."
                        For i = 1 To NumOfDecimals
                            Answer = Answer + "0"
                        Next i
                    End If

                'If decimal does exist, adjust the answer to
                'the number of decimal places that the user
                'specified
                Else
                    LenAfterDecimal = Len(Answer) - DecimalCheck
                    If LenAfterDecimal > NumOfDecimals Then
                        If NumOfDecimals = "0" Then
                            DecimalCheck = DecimalCheck - 1
                        End If
                        Answer = Mid(Answer, 1, DecimalCheck + NumOfDecimals)
                    Else
                        For i = 1 To (NumOfDecimals - LenAfterDecimal)
                            Answer = Answer + "0"
                        Next i
                    End If
                End If
            End If
        End If
        Answer = ">> " + Answer + Tag + vbNewLine + frmCalcSolver.txtAnswer.Text
    End If

    'Display final answer
    frmCalcSolver.txtAnswer.Text = Answer

    Exit Sub

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Sub

Private Sub ExtractToken()
Dim i As Integer

    '********************
    '* SCANNING ROUTINE *
    '********************

    'Set default values
    OutputString = ""
    OutputValue = 0
    ValueString = ""

    'If at the end of string, return EOS
    If CurrentEntryIndex > Len(InputString) Then
        OutputString = "EOS"
        Exit Sub
    End If

    'Get character to be examined
    Char = Mid(InputString, CurrentEntryIndex, 1)

    'Space
    If Char = " " Then
        CurrentEntryIndex = CurrentEntryIndex + 1
        ExtractToken
        Exit Sub
    End If

    'Operator or parenthesis
    If Char = "+" Or Char = "-" Or Char = "*" Or Char = "/" Or Char = "^" Or Char = "(" Or Char = ")" Or Char = "!" Then
        CurrentEntryIndex = CurrentEntryIndex + 1

        'Set return value
        OutputString = Char
        Exit Sub
    End If

    'Number
    If (Char >= "0" And Char <= "9") Or Char = "." Then

        'Digits before decimal
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Decimal
        While Char = "."
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Digits after decimal
        While Char >= "0" And Char <= "9"
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Set return values
        OutputString = "Number"
        OutputValue = CDbl(ValueString)
        Exit Sub
    End If

    'Return text language identifiers
    If LCase(Char) >= "a" And LCase(Char) <= "z" Then
        While (LCase(Char) >= "a" And LCase(Char) <= "z")
            ValueString = ValueString + Char
            CurrentEntryIndex = CurrentEntryIndex + 1
            If CurrentEntryIndex <= Len(InputString) Then
                Char = Mid(InputString, CurrentEntryIndex, 1)
            Else
                Char = ""
            End If
        Wend

        'Set return value
        OutputString = LCase(ValueString)
        Exit Sub
    End If

End Sub

Private Function GetE()
On Error GoTo ErrorHandler

    '*****************************
    '* PARSING ROUTINE (Level 1) *
    '*****************************

    'Get the lower value (T)
    Value = GetT

    'Exit function on error or help call
    If InError Or Help Then
        Exit Function
    End If

    Select Case OutputString

        'Addition operator
        Case "+"
            ExtractToken
            GetE = Value + GetE()

        'Subraction operator
        Case "-"
            ExtractToken
            GetE = Value - GetE()

        'Everything else passes upwards
        Case Else
            GetE = Value
    End Select

    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function GetT()
On Error GoTo ErrorHandler
Dim Exponent As Double

    '*****************************
    '* PARSING ROUTINE (Level 2) *
    '*****************************

    'Get the lower value (F)
    Value = GetF

    'Exit function on error or help call
    If InError Or Help Then
        Exit Function
    End If

    Select Case OutputString

        'Multiplication operator
        Case "*"
            ExtractToken
            GetT = Value * GetT()

        'Division operator
        Case "/"
            ExtractToken
            GetT = Value / GetT()

        Case Else
            GetT = Value
    End Select

    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function GetF()
On Error GoTo ErrorHandler
Dim Base As Double
Dim LogBase As String
Dim LogIndex As Long
Dim i As Integer

    '*****************************
    '* PARSING ROUTINE (Level 3) *
    '*****************************

    'Handle the low level calculations
    Select Case OutputString

        '***************
        'Basic Functions
        '***************

        'Number
        Case "Number"
            Value = OutputValue
            ExtractToken
            GetF = PostToken

        'Random number
        Case "rnd"
            Randomize
            Value = Rnd
            ExtractToken
            GetF = PostToken

        'Negative
        Case "-"
            ExtractToken
            Value = GetF()
            GetF = (-Value)

        'Parenthesis
        Case "("
            ExtractToken
            Value = GetE
            If OutputString <> ")" And OutputString <> "EOS" Then
                TrapErrors 0
                Exit Function
            End If
            If OutputString = ")" Then
                ExtractToken
                GetF = PostToken
            Else
                GetF = Value
            End If
            ExtractToken

        'Previous answer
        Case "ans"
            Value = PrevAnswer
            ExtractToken
            GetF = PostToken

        '*************
        'Miscellaneous
        '*************

        'Absolute value
        Case "abs"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                GetF = Abs(Value)
            End If

        'Help
        Case "help"
            Help = True

        'Square Root
        Case "sr"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                GetF = Sqr(Value)
            End If

        '*********
        'Constants
        '*********

        'e
        Case "e"
            GetF = Exp(1)
            ExtractToken

        'Pi
        Case "pi"
            GetF = Pi
            ExtractToken
            Exit Function

        '**********
        'Logarithms
        '**********

        'Logarithm (to a base)
        Case "log"

            'Get logarithm base
            LogBase = frmCalcSolver.txtLogBase.Text

            'If the box is empty, set it with the default 10
            If LogBase = "" Then
                frmCalcSolver.txtLogBase.Text = "10"
                Base = 10

            'Retrieve logarithm base
            Else
                Base = Val(LogBase)
            End If

            'Get number
            ExtractToken
            GetF = Log(GetBody) / Log(Base)

        'Natural logarithm
        Case "ln"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                GetF = Log(Value)
            End If

        '***********************
        'Trigonometric Functions
        '***********************

        'Cosine
        Case "cos"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = Cos(Value)
            End If

        'Cotangent
        Case "cot"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / Tan(Value)
            End If

        'Cosecant
        Case "csc"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / Sin(Value)
            End If

        'Hyperbolic cosecant
        Case "hcsc"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) - Exp(-Value))
            End If
            Exit Function

        'Hyperbolic cosine
        Case "hcos"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / 2
            End If

        'Hyperbolic cotangent
        Case "hcot"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) + Exp(-Value)) / (Exp(Value) - Exp(-Value))
            End If

        'Hyperbolic secant
        Case "hsec"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 2 / (Exp(Value) + Exp(-Value))
            End If

        'Hyperbolic sine
        Case "hsin"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / 2
            End If

        'Hyperbolic tangent
        Case "htan"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = (Exp(Value) - Exp(-Value)) / (Exp(Value) + Exp(-Value))
            End If

        'Inverse hyperbolic cosine
        Case "ihcos"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value - 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic cosecant
        Case "ihcsc"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log((Sgn(Value) * Sqr(Value * Value + 1) + 1) / Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic cotangent
        Case "ihcot"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log((Value + 1) / (Value - 1)) / 2
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic sine
        Case "ihsin"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log(Value + Sqr(Value * Value + 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic secant
        Case "ihsec"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log((Sqr(-Value * Value + 1) + 1) / Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse hyperbolic tangent
        Case "ihtan"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Log((1 + Value) / (1 - Value)) / 2
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cosecant
        Case "icsc"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + (Sgn(Value) - 1) * (2 * Atn(1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cosine
        Case "icos"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(-Value / Sqr(-Value * Value + 1)) + 2 * Atn(1)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse cotangent
        Case "icot"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(Value) + 2 * Atn(1)
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse secant
        Case "isec"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(Value * Value - 1)) + Sgn((Value) - 1) * (2 * Atn(1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse sine
        Case "isin"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(Value / Sqr(-Value * Value + 1))
                ConvertToDegrees
                GetF = Value
            End If

        'Inverse tangent
        Case "itan"
            ExtractToken
            Value = GetBody
            If InError Then
                Exit Function
            Else
                Value = Atn(Value)
                ConvertToDegrees
                GetF = Value
            End If

        'Secant
        Case "sec"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = 1 / Cos(Value)
            End If

        'Sine
        Case "sin"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = Sin(Value)
            End If

        'Tangent
        Case "tan"
            ExtractToken
            Value = GetBody
            ConvertToRadians
            If InError Then
                Exit Function
            Else
                GetF = Tan(Value)
            End If

        'Anything else is an error
        Case Else
            TrapErrors 0

    End Select

    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Function GetBody()
On Error GoTo ErrorHandler

    Select Case OutputString

        'Number
        Case "Number"
            Value = OutputValue
            ExtractToken
            GetBody = PostToken

        'Negative
        Case "-"
            ExtractToken
            Value = GetE
            GetBody = (-Value)
            ExtractToken

        'Random number
        Case "rnd"
            Randomize
            Value = Rnd
            ExtractToken
            GetBody = PostToken

        'Parenthesis
        Case "("
            ExtractToken
            Value = GetE
            If OutputString <> ")" And OutputString <> "EOS" Then
                TrapErrors 0
                Exit Function
            End If
            If OutputString = ")" Then
                ExtractToken
                GetBody = PostToken
            Else
                GetBody = Value
            End If
            ExtractToken

        'e
        Case "e"
            GetBody = Exp(1)
            ExtractToken

        'Pi
        Case "pi"
            GetBody = Pi
            ExtractToken
    End Select

    Exit Function

ErrorHandler:

    'Trap errors
    TrapErrors Err.Number

End Function

Private Sub TrapErrors(ErrNumber As Long)

    'Set trapped error message
    If ErrNumber = 6 Then
        'Overflow
        ErrorMessage = "Error: Overflow"
    ElseIf ErrNumber = 11 Then
        'Division By Zero
        ErrorMessage = "Error: Division By Zero"
    Else
        'Unknown error
        ErrorMessage = "Calculation Error"
    End If

    'Set return values
    InError = True
    OutputString = "TError"

End Sub

Private Sub ConvertToDegrees()

    'Convert to degrees
    If frmCalcSolver.optAngMode(0).Value = True Then
        Value = Value * (180 / Pi)
    End If

End Sub

Private Sub ConvertToRadians()

    'Convert to radians
    If frmCalcSolver.optAngMode(0).Value = True Then
        Value = Value * (Pi / 180)
    End If

End Sub

Private Function PostToken()
Dim Factorial As Double
Dim i As Integer

    'Ignore operators, EOS strings, and right parentheses
    If OutputString = "+" Or OutputString = "-" Or OutputString = "*" Or OutputString = "/" Or OutputString = "EOS" Or OutputString = ")" Then
        PostToken = Value

    'Handle special tokens that come after the value
    Else
        Select Case OutputString

            'Factorial
            Case "!"
                If (CDbl(Value) <> CLng(Value)) Or Value < 0 Then
                    TrapErrors 0
                    Exit Function
                End If
                Factorial = 1
                For i = Value To 1 Step -1
                    Factorial = Factorial * i
                Next i
                PostToken = Factorial
                ExtractToken

            'Exponent
            Case "^"
                ExtractToken
                PostToken = Value ^ GetF

            'Other "post" tokens multiply
            Case Else
                PostToken = Value * GetF
        End Select
    End If

End Function
