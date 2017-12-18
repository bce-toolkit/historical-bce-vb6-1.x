Attribute VB_Name = "mdlBalancer"
Option Explicit

Public Type tppFraction
    Numerator As Long
    Denominator As Long
End Type

Public Type tppElement
    Element As String
    Count As Long
End Type

Public Function BalanceCE(ByVal strEquation As String, ByRef lpResults() As Long) As Boolean
    On Error Resume Next
    Dim lpCurrent As Long
    Dim bpResults As Boolean
    Dim expEquations() As Long
    Dim expResults() As Long
    Dim lpMaxX As Long
    Dim lpMaxY As Long
    Dim expFractions() As tppFraction
    strEquation = RemoveSpace(strEquation)
    strEquation = ResolveStringCC(strEquation, "==", "=", "++", "+", "`", "", ".", "")
    bpResults = GetStack(strEquation, expEquations(), expResults(), lpMaxX, lpMaxY)
    If bpResults = False Then
        BalanceCE = False
        Exit Function
    End If
    bpResults = GetEquationResults(expEquations(), expResults(), expFractions(), lpMaxX, lpMaxY)
    If bpResults = False Then
        BalanceCE = False
        Exit Function
    End If
    FractionToNumber expFractions(), lpResults()
    For lpCurrent = 1 To UBound(lpResults())
        If lpResults(lpCurrent) <= 0 Then
            BalanceCE = False
            Exit Function
        End If
        DoEvents
    Next lpCurrent
    BalanceCE = True
End Function

Public Function RemoveSpace(ByVal strInput As String) As String
    On Error Resume Next
    Dim strTotal As String
    Dim lpCurrent As Long
    strTotal = vbNullString
    For lpCurrent = 1 To Len(strInput)
        If Mid(strInput, lpCurrent, 1) <> Space(1) Then
            strTotal = strTotal & Mid(strInput, lpCurrent, 1)
        End If
        DoEvents
    Next lpCurrent
    RemoveSpace = strTotal
End Function

Public Function StatisticsIn(ByVal strInput As String, ByVal strStatistics As String) As Long
    On Error Resume Next
    Dim lpCurrent As Long
    Dim strCurrent As String
    Dim lpStatistics As Long
    For lpCurrent = 1 To (Len(strInput) - Len(strStatistics) + 1)
        If Mid(strInput, lpCurrent, Len(strStatistics)) = strStatistics Then
            lpStatistics = lpStatistics + 1
        End If
    Next lpCurrent
    StatisticsIn = lpStatistics
End Function

Public Function ResolveStringCC(ByVal sString As String, ParamArray varReplacements() As Variant) As String
    On Error Resume Next
    Dim intMacro As Integer
    Dim strResString As String
    Dim strMacro As String
    Dim strValue As String
    Dim intPos As Integer
    strResString = sString
    For intMacro = LBound(varReplacements) To UBound(varReplacements) Step 2
        strMacro = varReplacements(intMacro)
        On Error GoTo MismatchedPairs
        strValue = varReplacements(intMacro + 1)
        On Error GoTo 0
        Do
            intPos = InStr(strResString, strMacro)
            If intPos > 0 Then
                strResString = Left$(strResString, intPos - 1) & strValue & Right$(strResString, Len(strResString) - Len(strMacro) - intPos + 1)
            End If
        Loop Until intPos = 0
    Next intMacro
    ResolveStringCC = strResString
    Exit Function
MismatchedPairs:
    Resume Next
End Function

Public Function ScanPrefixNumber(ByVal strInput As String, ByRef strVariant As String) As String
    On Error Resume Next
    Dim lpCurrent As Long
    Dim lpPosition As Long
    Dim strPrefix As String
    Dim strMiddle As String
    For lpCurrent = Len(strInput) To 1 Step -1
        strPrefix = Left(strInput, lpCurrent)
        If IsNumeric(strPrefix) = True Then
            strVariant = Right(strInput, Len(strInput) - lpCurrent)
            strMiddle = strPrefix
            Exit For
        End If
        DoEvents
    Next lpCurrent
    If Trim(strMiddle) = vbNullString Then
        ScanPrefixNumber = "1"
    Else
        ScanPrefixNumber = strMiddle
    End If
End Function

Public Function IsUpcaseChar(ByVal strChar As String) As Boolean
    On Error Resume Next
    If IsNumeric(strChar) = True Then
        IsUpcaseChar = False
        Exit Function
    End If
    IsUpcaseChar = IIf(strChar = UCase(strChar), True, False)
End Function

Public Function ResolveSingleElement(ByVal strString As String, ByRef nfsElements() As tppElement, ByVal lpCount As Long) As Boolean
    On Error Resume Next
    Dim lpScanNumber As Long
    Dim strMiddle As String
    Dim lpCurrent As Long
    Dim lpNextUpcase As Long
    Dim lpFinalPos As Long
    Dim strIncludeLeft As String
    Dim strIncludeRight As String
    Dim strCallback As String
    Dim strInput As String
    Dim lpNumber As Long
    Dim strReplace As String
    Dim lpPosition As Long
    Dim strPrefix As String
    Dim strVariant As String
    Dim lpP1 As Long, lpP2 As Long
    Dim lpN As Long
    Dim tmpElements() As tppElement
    Dim bpResult As Boolean
    Dim lpAdd As Long
    strInput = strString
    If StatisticsIn(strInput, "(") <> StatisticsIn(strInput, ")") Then
        ResolveSingleElement = False
        Exit Function
    End If
    Do
        lpPosition = InStr(1, strInput, "(")
        If lpPosition = 0 Then
            Exit Do
        End If
        lpP1 = 0
        lpP2 = 0
        For lpN = lpPosition To Len(strInput)
            Select Case Mid(strInput, lpN, 1)
                Case "("
                    lpP1 = lpP1 + 1
                Case ")"
                    lpP2 = lpP2 + 1
            End Select
            If lpP1 = lpP2 And lpP1 <> 0 Then
                Exit For
            End If
            lpFinalPos = lpN
            DoEvents
        Next lpN
        strIncludeLeft = Left(strInput, lpPosition - 1)
        strIncludeRight = Right(strInput, Len(strInput) - lpN)
        strReplace = Mid(strInput, lpPosition + 1, lpN - lpPosition - 1)
        bpResult = ResolveSingleElement(strReplace, tmpElements(), 0)
        If bpResult = False Then
            ResolveSingleElement = False
            Exit Function
        End If
        strPrefix = ScanPrefixNumber(strIncludeRight, strVariant)
        strIncludeRight = strVariant
        lpAdd = IIf(IsNumeric(strPrefix) = True, Val(strPrefix), "1")
        strCallback = vbNullString
        For lpCurrent = 1 To UBound(tmpElements())
            strCallback = strCallback & tmpElements(lpCurrent).Element & Trim(Str(tmpElements(lpCurrent).Count * lpAdd))
        Next lpCurrent
        strInput = strIncludeLeft & strCallback & strIncludeRight
        DoEvents
    Loop
    Do
        If Trim(strInput) = vbNullString Then
            Exit Do
        End If
        If IsUpcaseChar(Left(strInput, 1)) = False Then
            ResolveSingleElement = False
            Exit Function
        End If
        lpNextUpcase = Len(strInput) + 1
        For lpCurrent = 2 To Len(strInput)
            If IsUpcaseChar(Mid(strInput, lpCurrent, 1)) = True Then
                lpNextUpcase = lpCurrent
                Exit For
            End If
            DoEvents
        Next lpCurrent
        strMiddle = Left(strInput, lpNextUpcase - 1)
        lpNumber = 1
        If IsNumeric(Right(strMiddle, 1)) = True Then
            For lpScanNumber = 1 To Len(strMiddle)
                If IsNumeric(Right(strMiddle, Len(strMiddle) - lpScanNumber)) = True Then
                    lpNumber = Val(Right(strMiddle, Len(strMiddle) - lpScanNumber))
                    strMiddle = Left(strMiddle, lpScanNumber)
                    Exit For
                End If
                DoEvents
            Next lpScanNumber
        End If
        For lpCurrent = 1 To Len(strMiddle)
            If IsNumeric(Mid(strMiddle, lpCurrent, 1)) = True Then
                ResolveSingleElement = False
                Exit Function
            End If
            DoEvents
        Next lpCurrent
        For lpCurrent = 1 To lpCount
            If Trim(UCase(nfsElements(lpCurrent).Element)) = Trim(UCase(strMiddle)) Then
                nfsElements(lpCurrent).Count = nfsElements(lpCurrent).Count + lpNumber
                GoTo ExistReady
            End If
            DoEvents
        Next lpCurrent
        ReDim Preserve nfsElements(1 To (lpCount + 1)) As tppElement
        lpCount = lpCount + 1
        nfsElements(lpCount).Count = lpNumber
        nfsElements(lpCount).Element = strMiddle
ExistReady:
        If lpNextUpcase = Len(strInput) + 1 Then
            strInput = vbNullString
        Else
            strInput = Right(strInput, Len(strInput) - (lpNextUpcase - 1))
        End If
        DoEvents
    Loop
    ResolveSingleElement = True
End Function

Public Function MaxDivisor(ByVal x As Long, ByVal y As Long) As Long
    On Error Resume Next
    Dim isNegative As Boolean
    Dim lpNumber1 As Long
    Dim lpNumber2 As Long
    isNegative = IIf(x < 0 Or y < 0, True, False)
    lpNumber1 = Abs(x): lpNumber2 = Abs(y)
    If (lpNumber1 = 0) Or (lpNumber2 = 0) Then
        MaxDivisor = 1
        Exit Function
    End If
    While lpNumber1 <> lpNumber2
        If lpNumber1 > lpNumber2 Then
            lpNumber1 = lpNumber1 - lpNumber2
        ElseIf lpNumber1 < lpNumber2 Then
            lpNumber2 = lpNumber2 - lpNumber1
        End If
    Wend
    MaxDivisor = IIf(isNegative = True, -lpNumber1, lpNumber1)
End Function

Public Function MinMultiple(ByVal x As Long, ByVal y As Long) As Long
    On Error Resume Next
    Dim isNegative As Boolean
    Dim lpMax As Long
    Dim lpNumber1 As Long
    Dim lpNumber2 As Long
    isNegative = IIf(x < 0 Or y < 0, True, False)
    lpNumber1 = Abs(x): lpNumber2 = Abs(y)
    lpMax = MaxDivisor(lpNumber1, lpNumber2)
    If lpMax = 0 Then
        MinMultiple = 0
    Else
        MinMultiple = IIf(isNegative = True, -lpNumber1 * lpNumber2 / lpMax, lpNumber1 * lpNumber2 / lpMax)
    End If
End Function

Public Function FractionCreate(ByVal lpNumerator As Long, lpDenominator As Long) As tppFraction
    On Error Resume Next
    Dim lpsResult As tppFraction
    With lpsResult
        .Denominator = lpDenominator
        .Numerator = lpNumerator
    End With
    FractionCreate = lpsResult
End Function

Public Sub FractionSimplify(ByRef lpsVariable As tppFraction)
    On Error Resume Next
    Dim lpMaxDivisor As Long
    With lpsVariable
        lpMaxDivisor = MaxDivisor(.Numerator, .Denominator)
        .Numerator = .Numerator / lpMaxDivisor
        .Denominator = .Denominator / lpMaxDivisor
    End With
End Sub

Public Function FractionPlus(ByRef lpNumber1 As tppFraction, ByRef lpNumber2 As tppFraction) As tppFraction
    On Error Resume Next
    Dim lpsResult As tppFraction
    lpsResult = FractionCreate(lpNumber1.Numerator * lpNumber2.Denominator + lpNumber2.Numerator * lpNumber1.Denominator, lpNumber1.Denominator * lpNumber2.Denominator)
    FractionSimplify lpsResult
    FractionPlus = lpsResult
End Function

Public Function FractionMinus(ByRef lpNumber1 As tppFraction, ByRef lpNumber2 As tppFraction) As tppFraction
    On Error Resume Next
    Dim lpsResult As tppFraction
    lpsResult = FractionCreate(lpNumber1.Numerator * lpNumber2.Denominator - lpNumber2.Numerator * lpNumber1.Denominator, lpNumber1.Denominator * lpNumber2.Denominator)
    FractionSimplify lpsResult
    FractionMinus = lpsResult
End Function

Public Function FractionMultiplination(ByRef lpNumber1 As tppFraction, ByRef lpNumber2 As tppFraction) As tppFraction
    On Error Resume Next
    Dim lpsResult As tppFraction
    lpsResult = FractionCreate(lpNumber1.Numerator * lpNumber2.Numerator, lpNumber1.Denominator * lpNumber2.Denominator)
    FractionSimplify lpsResult
    FractionMultiplination = lpsResult
End Function

Public Function FractionDivision(ByRef lpNumber1 As tppFraction, ByRef lpNumber2 As tppFraction) As tppFraction
    On Error Resume Next
    Dim lpsResult As tppFraction
    lpsResult = FractionCreate(lpNumber1.Numerator * lpNumber2.Denominator, lpNumber1.Denominator * lpNumber2.Numerator)
    FractionSimplify lpsResult
    FractionDivision = lpsResult
End Function

Public Sub FractionToNumber(ByRef expResults() As tppFraction, ByRef lpResults() As Long)
    On Error Resume Next
    Dim lpMinMultiple As Long
    Dim lpCurrent As Long
    Dim lpLength As Long
    ReDim lpResults(1 To UBound(expResults())) As Long
    lpMinMultiple = 1
    lpLength = UBound(expResults())
    For lpCurrent = 1 To lpLength
        FractionSimplify expResults(lpCurrent)
    Next lpCurrent
    For lpCurrent = 1 To lpLength
        lpMinMultiple = MinMultiple(lpMinMultiple, expResults(lpCurrent).Denominator)
    Next lpCurrent
    For lpCurrent = 1 To lpLength
        lpResults(lpCurrent) = expResults(lpCurrent).Numerator * lpMinMultiple / expResults(lpCurrent).Denominator
    Next lpCurrent
End Sub

Public Function GetEquationResults(ByRef expEquation() As Long, ByRef expConstants() As Long, ByRef expResults() As tppFraction, ByVal lpMaxX As Long, ByVal lpMaxY As Long) As Boolean
    On Error Resume Next
    Dim lpCurrentX As Long, lpCurrentY As Long
    Dim lpCurrent As Long
    Dim lpTotalConstants As Long
    Dim lepTotal As tppFraction
    Dim lpTotalX As Long
    Dim expNew() As Long
    Dim expBuild() As Long
    Dim expBuildConstants() As Long
    Dim expNewConstants() As Long
    Dim vsAnswers() As tppFraction
    Dim lpNumber1 As Long, lpNumber2 As Long, lpNumber3 As Long
    Dim lpExchangeID As Long
    Dim bpResult As Boolean
    Dim lpMinMultiple As Long
    If lpMaxX > lpMaxY Then
        GetEquationResults = False
        Exit Function
    End If
    ReDim expResults(1 To lpMaxX) As tppFraction
    lpExchangeID = -1
    For lpCurrent = 1 To lpMaxY
        If expEquation(1, lpCurrent) <> 0 Then
            lpExchangeID = lpCurrent
            Exit For
        End If
        If lpCurrent = lpMaxY Then
            GetEquationResults = False
            Exit Function
        End If
        DoEvents
    Next lpCurrent
    If lpExchangeID = -1 Then
        GetEquationResults = False
        Exit Function
    End If
    For lpCurrentX = 1 To lpMaxX
        lpNumber1 = expEquation(lpCurrentX, 1)
        lpNumber2 = expEquation(lpCurrentX, lpExchangeID)
        expEquation(lpCurrentX, 1) = lpNumber2
        expEquation(lpCurrentX, lpExchangeID) = lpNumber1
        DoEvents
    Next lpCurrentX
    lpNumber1 = expConstants(1)
    lpNumber2 = expConstants(lpExchangeID)
    expConstants(1) = lpNumber2
    expConstants(lpExchangeID) = lpNumber1
    If lpMaxX = 1 Then
        lpNumber1 = 0: lpNumber2 = 0
        For lpCurrent = 1 To lpMaxY
            lpNumber1 = lpNumber1 + expEquation(1, lpCurrent)
            lpNumber2 = lpNumber2 + expConstants(lpCurrent)
            DoEvents
        Next lpCurrent
        expResults(1) = FractionCreate(lpNumber2, lpNumber1)
        FractionSimplify expResults(1)
        GetEquationResults = True
        Exit Function
    Else
        ReDim expBuild(1 To (lpMaxX - 1), 1 To lpMaxY) As Long
        ReDim expBuildConstants(1 To lpMaxY) As Long
        ReDim expNew(1 To (lpMaxX - 1), 1 To (lpMaxY - 1)) As Long
        ReDim expNewConstants(1 To (lpMaxY - 1)) As Long
        For lpCurrentY = 1 To lpMaxY
            If expEquation(1, lpCurrentY) = 0 Then
                For lpCurrentX = 2 To lpMaxX
                    expBuild(lpCurrentX - 1, lpCurrentY) = expEquation(lpCurrentX, lpCurrentY)
                    DoEvents
                Next lpCurrentX
                expBuildConstants(lpCurrentY) = expConstants(lpCurrentY)
            End If
            DoEvents
        Next lpCurrentY
        lpMinMultiple = 1
        For lpCurrentY = 1 To lpMaxY
            If expEquation(1, lpCurrentY) <> 0 Then
                lpMinMultiple = MinMultiple(lpMinMultiple, expEquation(1, lpCurrentY))
            End If
            DoEvents
        Next lpCurrentY
        For lpCurrentY = 1 To lpMaxY
            If expEquation(1, lpCurrentY) <> 0 Then
                For lpCurrentX = 2 To lpMaxX
                    expBuild(lpCurrentX - 1, lpCurrentY) = expEquation(lpCurrentX, lpCurrentY) * (lpMinMultiple / expEquation(1, lpCurrentY))
                Next lpCurrentX
                expBuildConstants(lpCurrentY) = expConstants(lpCurrentY) * (lpMinMultiple / expEquation(1, lpCurrentY))
            End If
            DoEvents
        Next lpCurrentY
        For lpCurrentY = 2 To lpMaxY
            If expEquation(1, lpCurrentY) <> 0 Then
                For lpCurrentX = 1 To (lpMaxX - 1)
                    expNew(lpCurrentX, lpCurrentY - 1) = expBuild(lpCurrentX, 1) - expBuild(lpCurrentX, lpCurrentY)
                Next lpCurrentX
                expNewConstants(lpCurrentY - 1) = expBuildConstants(1) - expBuildConstants(lpCurrentY)
            Else
                For lpCurrentX = 1 To (lpMaxX - 1)
                    expNew(lpCurrentX, lpCurrentY - 1) = expBuild(lpCurrentX, lpCurrentY)
                Next lpCurrentX
                expNewConstants(lpCurrentY - 1) = expBuildConstants(lpCurrentY)
            End If
            DoEvents
        Next lpCurrentY
        bpResult = GetEquationResults(expNew(), expNewConstants(), vsAnswers(), lpMaxX - 1, lpMaxY - 1)
        If bpResult = False Then
            GetEquationResults = False
            Exit Function
        End If
        lpTotalConstants = expConstants(1)
        lpTotalX = expEquation(1, 1)
        lepTotal = FractionCreate(0, 1)
        For lpCurrentX = 2 To lpMaxX
            lepTotal = FractionPlus(lepTotal, FractionMultiplination(FractionCreate(expEquation(lpCurrentX, 1), 1), vsAnswers(lpCurrentX - 1)))
            DoEvents
        Next lpCurrentX
        expResults(1) = FractionDivision(FractionMinus(FractionCreate(lpTotalConstants, 1), lepTotal), FractionCreate(lpTotalX, 1))
        For lpCurrent = 2 To lpMaxX
            expResults(lpCurrent) = vsAnswers(lpCurrent - 1)
        Next lpCurrent
        GetEquationResults = True
    End If
End Function

Public Function GetStack(ByVal strChemical As String, ByRef stkStack() As Long, ByRef stkConstants() As Long, ByRef lpMaxX As Long, ByRef lpMaxY As Long) As Boolean
    On Error Resume Next
    Dim psSides As New Collection
    Dim psSide1 As New Collection, psSide2 As New Collection
    Dim lpCurrent1 As Long, lpCurrent2 As Long, lpCurrent3 As Long
    Dim tmpSolves1() As tppElement, tmpSolves2() As tppElement, tmpSolves3() As tppElement, tmpSolves4() As tppElement
    Dim tmpLeft() As tppElement, tmpRight() As tppElement
    Dim lpSize As Long
    Dim lpCurrent As Long
    Dim bpResult As Boolean
    Dim bpAdded As Boolean
    Dim bpSwitch1 As Boolean
    Dim bpResolveEX As Boolean
    Dim lpCountEL As Long
    ClearCollection psSides
    ClearCollection psSide1
    ClearCollection psSide2
    ResolveCommandEX strChemical, psSides, "="
    If psSides.Count <> 2 Then
        GetStack = False
        Exit Function
    End If
    ResolveCommandEX psSides.Item(1), psSide1, "+"
    ResolveCommandEX psSides.Item(2), psSide2, "+"
    If Trim(psSide1.Item(1)) = vbNullString Or Trim(psSide1.Item(1)) = vbNullString Then
        GetStack = False
        Exit Function
    End If
    For lpCurrent = 1 To psSide1.Count
        bpResult = ResolveSingleElement(psSide1.Item(lpCurrent), tmpSolves2(), 0)
        If bpResult = False Then
            GetStack = False
            Exit Function
        End If
        If lpCurrent = 1 Then
            CopyElementStack tmpSolves2(), tmpSolves1()
            If psSide1.Count = 1 Then
                CopyElementStack tmpSolves1(), tmpSolves3()
            End If
        Else
            For lpCurrent1 = 1 To UBound(tmpSolves2())
                lpSize = UBound(tmpSolves1()) + 1
                ReDim Preserve tmpSolves1(1 To lpSize) As tppElement
                tmpSolves1(lpSize) = tmpSolves2(lpCurrent1)
            Next lpCurrent1
            lpSize = 0
            For lpCurrent1 = 1 To UBound(tmpSolves1())
                bpAdded = False
                For lpCurrent2 = 1 To lpSize
                    If Trim(UCase(tmpSolves3(lpCurrent2).Element)) = Trim(UCase(tmpSolves1(lpCurrent1).Element)) Then
                        tmpSolves3(lpCurrent2).Count = tmpSolves3(lpCurrent2).Count + tmpSolves1(lpCurrent1).Count
                        bpAdded = True
                    End If
                Next lpCurrent2
                If bpAdded = False Then
                    lpSize = lpSize + 1
                    ReDim Preserve tmpSolves3(1 To lpSize) As tppElement
                    tmpSolves3(lpSize) = tmpSolves1(lpCurrent1)
                End If
            Next lpCurrent1
        End If
        DoEvents
    Next lpCurrent
    CopyElementStack tmpSolves3(), tmpLeft()
    For lpCurrent = 1 To psSide2.Count
        bpResult = ResolveSingleElement(psSide2.Item(lpCurrent), tmpSolves2(), 0)
        If bpResult = False Then
            GetStack = False
            Exit Function
        End If
        If lpCurrent = 1 Then
            CopyElementStack tmpSolves2(), tmpSolves1()
            If psSide2.Count = 1 Then
                CopyElementStack tmpSolves1(), tmpSolves3()
            End If
        Else
            For lpCurrent1 = 1 To UBound(tmpSolves2())
                lpSize = UBound(tmpSolves1()) + 1
                ReDim Preserve tmpSolves1(1 To lpSize) As tppElement
                tmpSolves1(lpSize) = tmpSolves2(lpCurrent1)
            Next lpCurrent1
            lpSize = 0
            For lpCurrent1 = 1 To UBound(tmpSolves1())
                bpAdded = False
                For lpCurrent2 = 1 To lpSize
                    If Trim(UCase(tmpSolves3(lpCurrent2).Element)) = Trim(UCase(tmpSolves1(lpCurrent1).Element)) Then
                        tmpSolves3(lpCurrent2).Count = tmpSolves3(lpCurrent2).Count + tmpSolves1(lpCurrent1).Count
                        bpAdded = True
                    End If
                Next lpCurrent2
                If bpAdded = False Then
                    lpSize = lpSize + 1
                    ReDim Preserve tmpSolves3(1 To lpSize) As tppElement
                    tmpSolves3(lpSize) = tmpSolves1(lpCurrent1)
                End If
            Next lpCurrent1
        End If
        DoEvents
    Next lpCurrent
    CopyElementStack tmpSolves3(), tmpRight()
    For lpCurrent1 = 1 To UBound(tmpLeft())
        bpSwitch1 = False
        For lpCurrent2 = 1 To UBound(tmpRight())
            If Trim(UCase(tmpLeft(lpCurrent1).Element)) = Trim(UCase(tmpRight(lpCurrent2).Element)) Then
                bpSwitch1 = True
                Exit For
            End If
            DoEvents
        Next lpCurrent2
        If bpSwitch1 = False Then
            GetStack = False
            Exit Function
        End If
        DoEvents
    Next lpCurrent1
    For lpCurrent1 = 1 To UBound(tmpRight())
        bpSwitch1 = False
        For lpCurrent2 = 1 To UBound(tmpLeft())
            If Trim(UCase(tmpLeft(lpCurrent2).Element)) = Trim(UCase(tmpRight(lpCurrent1).Element)) Then
                bpSwitch1 = True
                Exit For
            End If
            DoEvents
        Next lpCurrent2
        If bpSwitch1 = False Then
            GetStack = False
            Exit Function
        End If
        DoEvents
    Next lpCurrent1
    lpMaxX = psSide1.Count + psSide2.Count
    lpMaxY = UBound(tmpLeft()) + 1
    ReDim stkStack(1 To lpMaxX, 1 To lpMaxY) As Long
    ReDim stkConstants(1 To lpMaxY) As Long
    For lpCurrent1 = 1 To (lpMaxY - 1)
        For lpCurrent2 = 1 To psSide1.Count
            bpResolveEX = ResolveSingleElement(psSide1.Item(lpCurrent2), tmpSolves4(), 0)
            If bpResolveEX = False Then
                GetStack = False
                Exit Function
            End If
            For lpCurrent3 = 1 To UBound(tmpSolves4())
                If Trim(UCase(tmpSolves4(lpCurrent3).Element)) = Trim(UCase(tmpLeft(lpCurrent1).Element)) Then
                    stkStack(lpCurrent2, lpCurrent1) = tmpSolves4(lpCurrent3).Count
                    Exit For
                End If
            Next lpCurrent3
            DoEvents
        Next lpCurrent2
        For lpCurrent2 = 1 To psSide2.Count
            bpResolveEX = ResolveSingleElement(psSide2.Item(lpCurrent2), tmpSolves4(), 0)
            If bpResolveEX = False Then
                GetStack = False
                Exit Function
            End If
            For lpCurrent3 = 1 To UBound(tmpSolves4())
                If Trim(UCase(tmpSolves4(lpCurrent3).Element)) = Trim(UCase(tmpLeft(lpCurrent1).Element)) Then
                    stkStack(psSide1.Count + lpCurrent2, lpCurrent1) = -tmpSolves4(lpCurrent3).Count
                    Exit For
                End If
            Next lpCurrent3
            DoEvents
        Next lpCurrent2
        DoEvents
    Next lpCurrent1
    stkStack(1, lpMaxY) = 1
    stkConstants(lpMaxY) = 1
    GetStack = True
End Function

Public Sub CopyElementStack(ByRef Stack1() As tppElement, ByRef Stack2() As tppElement)
    On Error Resume Next
    Dim lpCurrent As Long
    ReDim Stack2(1 To UBound(Stack1())) As tppElement
    For lpCurrent = 1 To UBound(Stack1())
        Stack2(lpCurrent) = Stack1(lpCurrent)
    Next lpCurrent
End Sub

Public Sub ClearCollection(ByRef bCollection As Collection)
    On Error Resume Next
    Dim lCurrent As Long
    For lCurrent = bCollection.Count To 1 Step -1
        bCollection.Remove lCurrent
        DoEvents
    Next lCurrent
End Sub

Public Sub ResolveCommandEX(ByVal sCommand As String, ByRef sSliced As Collection, Optional ByVal strChar As String = "+")
    On Error Resume Next
    Dim sLocal As String
    Dim lPrevious As Long
    Dim lCurrent As Long
    Dim lFirst As Long
    Dim sTemporary As String
    lPrevious = 1
    sLocal = Trim(sCommand)
ReRouteString:
    If Left(sLocal, 1) = strChar Then
        sLocal = Right(sLocal, Len(sLocal) - 1)
        GoTo ReRouteString
    End If
    If Right(sLocal, 1) = strChar Then
        sLocal = Left(sLocal, Len(sLocal) - 1)
        GoTo ReRouteString
    End If
    Do
        lFirst = InStr(1, sLocal, strChar)
        If lFirst = 0 Then
            sSliced.Add sLocal
            Exit Do
        End If
        sTemporary = Mid(sLocal, lPrevious, lFirst - lPrevious)
        sSliced.Add sTemporary
        sLocal = Right(sLocal, Len(sLocal) - lFirst)
        DoEvents
    Loop
    DoEvents
End Sub
