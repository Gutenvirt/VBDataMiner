Public Class DMDataAnalyzer


    Dim itemPValue() As Single
    Dim stuRawScore() As Single


    Dim numScorableItems() As Integer
    Dim StuRawScoreBins(10) As Integer


    Dim colSkip As Integer = -1

    Dim testMedian As Integer
    Dim testQ1 As Integer
    Dim testQ3 As Integer

    Dim testSTDev As Single
    Dim testMean As Single
    Dim testAlpha As Single
    Dim testSTDEM As Single
    Dim testSkewness As Single

    Dim LoadedData As New DMDataSet


    Public Sub CreateBinaryDataField()
        

    End Sub

    Public Sub CalculatePValue()

        Dim _t As String
        Dim _x As Integer = 0

        ReDim itemPValue(numVar)
        ReDim numScorableItems(numVar)

        Dim _tPvCR(numObs) As Integer


        For _a = 0 To numVar - 1
            For _b = 0 To numObs - 1
                _t = BinData(_a, _b)

                If _t <> NullValue And colSkip <> _a Then
                    If itemType(_a) = "CR" Then
                        numScorableItems(_a) = 1
                        _tPvCR(_b) = CInt(RawData(_a, _b).Replace("+", ""))
                    Else
                        itemPValue(_a) += CInt(_t)
                        numScorableItems(_a) += 1
                    End If

                End If
            Next
            If itemType(_a) = "CR" Then itemPValue(_a) = _tPvCR.Average / (_tPvCR.Max - _tPvCR.Min)

        Next

        For _a = 0 To numVar - 1
            If colSkip <> _a And itemType(_a) <> "CR" Then itemPValue(_a) = itemPValue(_a) / numScorableItems(_a)
        Next

    End Sub

    Public Sub CalculateSTDev()
        ReDim stuRawScore(numObs)

        Dim _sumScoreVariance2 As Double

        For _a = 0 To numObs - 1
            For _b = 0 To numVar - 1
                If BinData(_b, _a) = "1" And colSkip <> _b Then stuRawScore(_a) += CInt(BinData(_b, _a))
            Next
        Next

        testMean = stuRawScore.Average

        For _a = 0 To numObs - 1
            _sumScoreVariance2 += Math.Pow(stuRawScore(_a) - testMean, 2)
        Next

        testSTDev = Math.Sqrt(_sumScoreVariance2 / numObs)
    End Sub

    Public Sub CalculateAlpha()
        Dim _sumPVariance As Single = 0

        For _a = 0 To numVar - 1
            If colSkip <> _a Then _sumPVariance += itemPValue(_a) * (1 - itemPValue(_a))
        Next
        testAlpha = (numVar / (numVar - 1)) * ((Math.Pow(testSTDev, 2) - _sumPVariance) / Math.Pow(testSTDev, 2))
        testSTDEM = testSTDev * Math.Sqrt(1 - testAlpha)
    End Sub

    Public Sub CalculatePointBiSerial()

        ReDim itemPointBiserial(numVar)

        Dim _meanCorrect(numVar)
        Dim _meanWrong(numVar)

        Dim _numStuCorrect(numVar)
        Dim _numStuWrong(numVar)

        For _a = 0 To numVar - 1
            For _b = 0 To numObs - 1
                If itemType(_a) <> "CR" Then
                    Select Case BinData(_a, _b)
                        Case "1"
                            _meanCorrect(_a) += stuRawScore(_b)
                            _numStuCorrect(_a) += 1
                        Case "0"
                            _meanWrong(_a) += stuRawScore(_b)
                            _numStuWrong(_a) += 1
                        Case "NaN"
                            _meanWrong(_a) += stuRawScore(_b)
                            _numStuWrong(_a) += 1
                    End Select
                Else
                    Exit For
                End If
            Next
        Next

        For _a = 0 To numVar - 1
            _meanCorrect(_a) = _meanCorrect(_a) / _numStuCorrect(_a)
            _meanWrong(_a) = _meanWrong(_a) / _numStuWrong(_a)
            itemPointBiserial(_a) = (_meanCorrect(_a) - _meanWrong(_a)) / testSTDev * Math.Sqrt(_numStuCorrect(_a) / numObs * _numStuWrong(_a) / numObs)
        Next

    End Sub

    Public Sub CalculateAlphaIfDropped()

        ReDim testAlphaDrop(numVar)

        For _h = 0 To numVar - 1

            colSkip = _h

            CalculatePValue()
            CalculateSTDev()
            CalculateAlpha()

            testAlphaDrop(_h) = testAlpha

        Next

        colSkip = -1

        CalculatePValue()
        CalculateSTDev()
        CalculateAlpha()

    End Sub

    Public Sub CalculateDescriptiveStats(StudentScores() As Single)
        Array.Sort(StudentScores)
        If StudentScores.Length Mod 2 <> 0 Then
            testMedian = StudentScores(StudentScores.GetUpperBound(0) / 2)
            testQ1 = StudentScores(StudentScores.GetUpperBound(0) / 4)
            testQ3 = StudentScores(3 * StudentScores.GetUpperBound(0) / 4)
        Else
            testMedian = (StudentScores(StudentScores.Length \ 2) + StudentScores((StudentScores.Length \ 2) - 1)) \ 2
            testQ1 = (StudentScores(StudentScores.Length \ 4) + StudentScores((StudentScores.Length \ 4) - 1)) \ 2
            testQ3 = (StudentScores(3 * StudentScores.Length \ 4) + StudentScores(3 * (StudentScores.Length \ 4) - 1)) \ 2
        End If

        testSkewness = 3 * (testMean - testMedian) / testSTDev

        Dim _s As Single

        For m = 0 To StudentScores.Length - 1

            _s = StudentScores(m) / numVar

            If _s < 0.1 Then StuRawScoreBins(0) += 1
            If _s >= 0.1 And _s < 0.2 Then StuRawScoreBins(1) += 1
            If _s >= 0.2 And _s < 0.3 Then StuRawScoreBins(2) += 1
            If _s >= 0.3 And _s < 0.4 Then StuRawScoreBins(3) += 1
            If _s >= 0.4 And _s < 0.5 Then StuRawScoreBins(4) += 1
            If _s >= 0.5 And _s < 0.6 Then StuRawScoreBins(5) += 1
            If _s >= 0.6 And _s < 0.7 Then StuRawScoreBins(6) += 1
            If _s >= 0.7 And _s < 0.8 Then StuRawScoreBins(7) += 1
            If _s >= 0.8 And _s < 0.9 Then StuRawScoreBins(8) += 1
            If _s >= 0.9 Then StuRawScoreBins(9) += 1

        Next
    End Sub
End Class
