Option Explicit On

Imports System.Data
Imports System.Data.DataSet
Imports System.IO

Public Class frmMain
    Dim myStreamWriter As System.IO.StreamWriter

    Dim dData As System.Data.DataSet

    Dim RawData(,) As String
    Dim BinData(,) As String

    Dim DataName As String = ""
    Dim NullValue As String = "NaN"
    Dim itemType() As String
    Dim ansKey() As String

    Dim itemPValue() As Single
    Dim stuRawScore() As Single
    Dim itemOmission() As Single
    Dim itemChoiceFreq_1() As Single
    Dim itemChoiceFreq_2() As Single
    Dim itemChoiceFreq_3() As Single
    Dim itemChoiceFreq_4() As Single
    Dim itemPointBiserial() As Single
    Dim testAlphaDrop() As Single

    Dim MCCount() As Integer
    Dim MSCount() As Integer
    Dim CRCount() As Integer
    Dim GRCount() As Integer

    Dim numScorableItems() As Integer
    Dim StuRawScoreBins(10) As Integer

    Dim numVar As Integer
    Dim numObs As Integer
    Dim numStuDrop As Integer

    Dim colSkip As Integer = -1

    Dim testMedian As Integer
    Dim testQ1 As Integer
    Dim testQ3 As Integer

    Dim testSTDev As Single
    Dim testMean As Single
    Dim testAlpha As Single
    Dim testSTDEM As Single
    Dim testSkewness As Single


    Public Sub CreateBinaryDataField()
        numVar = dData.Tables(0).Columns.Count
        numObs = dData.Tables(0).Rows.Count

        ReDim RawData(numVar, numObs)
        ReDim BinData(numVar, numObs)

        Dim _t As String

        ReDim itemChoiceFreq_1(numVar)
        ReDim itemChoiceFreq_2(numVar)
        ReDim itemChoiceFreq_3(numVar)
        ReDim itemChoiceFreq_4(numVar)
        ReDim itemOmission(numVar)

        ReDim itemType(numVar)
        ReDim MCCount(numVar)
        ReDim MSCount(numVar)
        ReDim CRCount(numVar)
        ReDim GRCount(numVar)

        ReDim ansKey(numVar)

        For _a = 0 To numVar - 1
            For _b = 0 To numObs - 1
                If dData.Tables(0).Rows(_b).Item(_a).ToString() = DBNull.Value.ToString Or dData.Tables(0).Rows(_b).Item(_a).ToString() = "" Then
                    RawData(_a, _b) = NullValue
                    BinData(_a, _b) = NullValue

                    itemOmission(_a) += 1
                Else
                    _t = dData.Tables(0).Rows(_b).Item(_a)

                    RawData(_a, _b) = _t
                    If _t.IndexOf("+") > -1 Then
                        ansKey(_a) = _t.Replace("+", "")
                        BinData(_a, _b) = "1"
                    Else
                        BinData(_a, _b) = "0"
                    End If

                    Select Case _t.Replace("+", "")
                        Case "A", "F"
                            itemChoiceFreq_1(_a) += 1
                            MCCount(_a) += 1
                        Case "B", "G"
                            itemChoiceFreq_2(_a) += 1
                            MCCount(_a) += 1
                        Case "C", "H"
                            itemChoiceFreq_3(_a) += 1
                            MCCount(_a) += 1
                        Case "D", "J"
                            itemChoiceFreq_4(_a) += 1
                            MCCount(_a) += 1
                    End Select

                    If _t.IndexOf(",") > -1 Then
                        MSCount(_a) += 1
                    Else
                        If _t.IndexOf("+") = 0 And IsNumeric(_t.Replace("+", "")) = True Then
                            If _t.Length = 2 Then CRCount(_a) += 1
                        Else
                            If IsNumeric(_t.Replace("+", "")) = True Then
                                GRCount(_a) += 1
                            End If
                        End If
                    End If

                End If
            Next
        Next

        dData.Dispose()

        For _a = 0 To numVar - 1

            Select Case Math.Max(Math.Max(MCCount(_a), CRCount(_a)), Math.Max(MSCount(_a), GRCount(_a)))
                Case MCCount(_a)
                    itemType(_a) = "MC"
                Case MSCount(_a)
                    itemType(_a) = "MS"
                Case CRCount(_a)
                    itemType(_a) = "CR"
                Case GRCount(_a)
                    itemType(_a) = "GR"
            End Select

            ansKey(_a) = ansKey(_a).Replace("A", "1")
            ansKey(_a) = ansKey(_a).Replace("B", "2")
            ansKey(_a) = ansKey(_a).Replace("C", "3")
            ansKey(_a) = ansKey(_a).Replace("D", "4")
            ansKey(_a) = ansKey(_a).Replace("F", "1")
            ansKey(_a) = ansKey(_a).Replace("G", "2")
            ansKey(_a) = ansKey(_a).Replace("H", "3")
            ansKey(_a) = ansKey(_a).Replace("J", "4")
        Next

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

    Public Sub ReadDataFile(FileName As String)

        Dim MyConnection As System.Data.OleDb.OleDbConnection
        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter


        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & FileName & "'; Extended Properties='Excel 12.0;IMEX=1;HDR=NO;Empty Text Mode=NullAsEmpty'")
        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$]", MyConnection)
        MyCommand.TableMappings.Add("Table", "RawData")

        dData = New System.Data.DataSet
        dData.Tables.Add("RawData")

        MyCommand.Fill(dData.Tables("RawData"))

        MyConnection.Close()

        DataName = dData.Tables(0).Rows(0).Item(6).ToString
        DataSetCleanup()

        Dim _numObs As Integer = dData.Tables("RawData").Rows.Count
        Dim _numVar As Integer = dData.Tables("RawData").Columns.Count

        Dim _bNull As Byte = 0
        Dim _nullRows As New List(Of Integer)


        For _a = 0 To _numObs - 1
            _bNull = 0
            For _b = 0 To _numVar - 1
                If dData.Tables("RawData").Rows(_a).IsNull(_b) Then
                    _bNull += 1
                Else
                    If dData.Tables("RawData").Rows(_a).Item(_b) = "" Then _bNull += 1
                End If

            Next
            If _bNull = _numVar Then _nullRows.Add(_a)
        Next

        For Each nullRow In _nullRows
            dData.Tables("RawData").Rows.RemoveAt(nullRow)
            numStuDrop += 1
        Next

        testSTDev = 0
        testMean = 0
        testMedian = 0
        testAlpha = 0
        testSTDEM = 0
        testSkewness = 0

    End Sub


    Public Sub DataSetCleanup()
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(0)
        dData.Tables(0).Columns.RemoveAt(dData.Tables(0).Columns.Count - 1)
        dData.Tables(0).Columns.RemoveAt(dData.Tables(0).Columns.Count - 1)
        dData.Tables(0).Columns.RemoveAt(dData.Tables(0).Columns.Count - 1)
        dData.Tables(0).Columns.RemoveAt(dData.Tables(0).Columns.Count - 1)
        dData.Tables(0).Columns.RemoveAt(dData.Tables(0).Columns.Count - 1)
        dData.Tables(0).Rows.RemoveAt(0)
        dData.Tables(0).Rows.RemoveAt(0)
        dData.Tables(0).Rows.RemoveAt(0)
        dData.Tables(0).Rows.RemoveAt(0)
        dData.Tables(0).Rows.RemoveAt(0)
    End Sub

    Public Sub HTMLOut()

        Dim strHTML As String

        strHTML = "<!DOCTYPE html><HTML><HEAD><TITLE>Test Statisics</TITLE>"
        strHTML &= "<STYLE>" & _
            ".graph{position: relative; width: 495px; height: 250px;}" & _
            ".bar{width: 45px; margin: 1px; display: inline-block; border: 1px solid #c1c1c1; position: relative; background-color: #edf2f9; vertical-align: baseline;}" & _
            ".median{top: 1px ;left: " & CInt(testMean / numVar * 495) & "px; height: 231px; width: 0px; margin: 0px; display: inline-block; border: 1px solid #9F9F9F; position: absolute; background-color: #FBE2E0; vertical-align: baseline;}" & _
            ".std{left: " & CInt((testMean - testSTDev) / numVar * 495) & "px; top: 234px; width: " & CInt(testSTDev / numVar * 2 * 495) & "px; margin: 0px; display: inline-block; border: 1px solid #9F9F9F; position: absolute; vertical-align: baseline;}" & _
            ".percent100{top: 0px; left: 0px; text-align: left; position: absolute; display: inline-block;font-family: Arial, Helvetica, Helv;font-size: x-small;font-style: normal;font-weight: normal;} .tablecenter{margin-left: auto; margin-right: auto;}" & _
            ".percent75{top: 58px; left: 0px; text-align: left; position: absolute; display: inline-block;font-family: Arial, Helvetica, Helv;font-size: x-small;font-style: normal;font-weight: normal;}" & _
            ".percent50{top: 115px; left: 0px; text-align: left; position: absolute; display: inline-block;font-family: Arial, Helvetica, Helv;font-size: x-small;font-style: normal;font-weight: normal;}" & _
            ".percent25{top: 173px; left: 0px; text-align: left; position: absolute; display: inline-block;font-family: Arial, Helvetica, Helv;font-size: x-small;font-style: normal;font-weight: normal;}" & _
            ".xlabel{text-align: center; width: 49px; margin: 0px; display: inline-block; position: relative; vertical-align: baseline; font-family: Arial, Helvetica, Helv;font-size: x-small;font-style: normal;font-weight: normal;}" & _
            "td{background-color: #ffffff;border-color: #c1c1c1;border-style: solid;border-width: 0 1px 1px 0;font-family: Arial, Helvetica, Helv;font-size: small;font-style: normal;font-weight: normal;text-align: right;vertical-align: middle;padding: 3px 6px;}" & _
            "th{background-color: #edf2f9;border-color: #b0b7bb;border-style: solid;border-width: 0 1px 1px 0;color: #112277;font-family: Arial, Helvetica, Helv;font-size: small;font-style: normal;font-weight: bold;text-align: center;vertical-align: middle;padding: 3px 6px;}" & _
            "table{border-color: #c1c1c1;border-style: solid;border-width: 1px 0 0 1px;border-collapse: collapse;border-spacing: 0;vertical-align: middle;}" & _
            ".inline{display: float; padding: 3px; border-width: 0px; border-color: #ffffff;} .center{text-align: center;} .left{text-align: left}</STYLE>"

        strHTML &= "</HEAD><BODY><TABLE class=""tablecenter""><tr><th colspan=""6""><p>" & DataName.ToUpper & "</p></th><th colspan=""4""><p>DataMiner CTT Report</p></th></tr><tr><td class=""center"" colspan=""3"">Printed on: " & Date.Today & "</td><td class=""center"" colspan=""3"">Generated by: " & Environment.UserName & "</td><td class=""center"" colspan=""4"">2014-15 by Chris Stefancik</td></tr>"

        strHTML &= "<tr><th colspan=""2""><p>Raw Score Statistics</p></th><th colspan=""8""><p>Percent Score Distribution</p></th></tr>"
        strHTML &= "<tr><td>Items</td><td>" & numVar & "</td>"
        strHTML &= "<td ROWSPAN=""12"" colspan=""8""><div><div class=""graph"">"
        For _a = 0 To 9
            strHTML &= "<div style=""height: " & CInt(StuRawScoreBins(_a) / StuRawScoreBins.Max * 230) & "px""; class=""bar""></div>"
        Next
        strHTML &= "<div class=""std""></div>"
        strHTML &= "<div class=""median""></div>"
        strHTML &= "<div class=""percent100"">" & CInt(StuRawScoreBins.Max / numObs * 100) & "%</div>"
        strHTML &= "<div class=""percent75"">" & CInt(StuRawScoreBins.Max / numObs * 75) & "%</div>"
        strHTML &= "<div class=""percent50"">" & CInt(StuRawScoreBins.Max / numObs * 50) & "%</div>"
        strHTML &= "<div class=""percent25"">" & CInt(StuRawScoreBins.Max / numObs * 25) & "%</div>"

        For _a = 0 To 9
            If _a = 9 Then
                strHTML &= "<div class=""xlabel"">" & _a * 10 & "-100</div>"
            Else
                strHTML &= "<div class=""xlabel"">" & _a * 10 & "-" & _a * 10 + 9 & "</div>"
            End If
        Next
        strHTML &= "</div></div></tr>"
        strHTML &= "<tr><td>Students</td><td>" & numObs & "</td></tr>"
        strHTML &= "<tr><td>Min</td><td>" & stuRawScore.Min & "</td></tr>"
        strHTML &= "<tr><td>Mean</td><td>" & testMean.ToString("0.00") & "</td></tr>"
        strHTML &= "<tr><td>Max</td><td>" & stuRawScore.Max & "</td></tr>"
        strHTML &= "<tr><td>Std Dev</td><td>" & testSTDev.ToString("0.00") & "</td></tr>"
        strHTML &= "<tr><td>Q1</td><td>" & testQ1 & "</td></tr>"
        strHTML &= "<tr><td>Median</td><td>" & testMedian & "</td></tr>"
        strHTML &= "<tr><td>Q3</td><td>" & testQ3 & "</td></tr>"
        strHTML &= "<tr><td>Skew</td><td>" & testSkewness.ToString("0.00") & "</td></tr>"
        strHTML &= "<tr><td>Alpha</td><td>" & testAlpha.ToString("0.00") & "</td></tr>"
        strHTML &= "<tr><td>SEM</td><td>" & testSTDEM.ToString("0.00") & "</td></tr>"

        If CInt((testAlpha - 1) * 0.8 / testAlpha / (-0.2) * numVar) - numVar > 0 Then strHTML &= "<tr><td class=""left"" colspan=""10"">Note: Adding " & CInt((testAlpha - 1) * 0.8 / testAlpha / (-0.2) * numVar) - numVar & " more similar questions could raise test reliability up to 0.80</td></tr>"
        If numStuDrop > 0 Then strHTML &= "<tr><td class=""left"" colspan=""10"">Note: Some students (" & numStuDrop & " total) were dropped because of complete test omission. </td></tr>"
        If CRCount.Max > 0 Then strHTML &= "<tr><td class=""left"" colspan=""10"">Note: This test contains constructed response items; the P-values has been calculated accordingly. </td></tr>"


        strHTML &= "<tr><th colspan=""4""><p>Item Statistics</p></th><th colspan=""6""><p>Selection Statistics</p></th></tr>"

        strHTML &= "<tr><th>Item</th><th>P-Value</th><th>PBS</th><th>Alpha Without</th><th>Answer</th><th>% C1</th><th>% C2</th><th>% C3</th><th>% C4</th><th> % Omitted</th></tr>"

        For _a = 0 To numVar - 1

            strHTML &= "<tr><td>" & _a + 1 & " " & itemType(_a) & "</td><td>" & _
                itemPValue(_a).ToString("0.00") & "</td><td>" & _
                itemPointBiserial(_a).ToString("0.00") & "</td><td>" & _
                testAlphaDrop(_a).ToString("0.00") & "</td>"
            strHTML &= "<td class=""center"">" & ansKey(_a) & "</td><td>"
            strHTML &= Math.Round(itemChoiceFreq_1(_a) / numObs * 100, 0) & "</td><td>" & _
                    Math.Round(itemChoiceFreq_2(_a) / numObs * 100, 0) & "</td><td>" & _
                    Math.Round(itemChoiceFreq_3(_a) / numObs * 100, 0) & "</td><td>" & _
                    Math.Round(itemChoiceFreq_4(_a) / numObs * 100, 0) & "</td><td>" & _
                    Math.Round(itemOmission(_a) / numObs * 100, 1) & "</td></tr>"
        Next
        'strHTML &= "<tr><th colspan=""10"">Citations</th></tr><tr><td colspan=""10"">" & _
        '"Afifi, A. A., & Elashoff, R. M. (1966). Missing observations in multivariate statistics I. Review of the literature. <em>Journal of the American Statistical Association, </em> 61(315), 595-604.<br></br>" & _
        '"Brown, J. D. (2001). Point-biserial correlation coefficients. <em>JALT Testing & Evaluation SIG Newsletter, </em> 5(3), 12-15.<br></br>" & _
        '"Brown, S. (2011). Measures of shape: Skewness and Kurtosis. Retrieved on December, 31, 2014.<br></br>" & _
        '"Ebel, R. L. (1950). Construction and validation of educational tests. <em>Review of Educational Research,</em> 87-97.<br></br>" & _
        '"Ebel, R. L. (1965). Confidence Weighting and Test Reliability. <em>Journal of Educational Measurement,</em> 2(1), 49-57.<br></br>" & _
        '"Kelley, T., Ebel, R., & Linacre, J. M. (2002). Item discrimination indices. <em>Rasch Measurement Transactions,</em> 16(3), 883-884.<br></br>" & _
        '"Krishnan, V. (2013). The Early Child Development Instrument (EDI): An item analysis using Classical Test Theory (CTT) on Alberta’s data. <em>Early Child Mapping (ECMap) Project Alberta, Community-University Partnership (CUP), Faculty of Extension, University of Alberta, Edmonton, Alberta.</em><br></br>" & _
        '"Matlock-Hetzel, S. (1997). Basic Concepts in Item and Test Analysis.<br></br>Pearson, K. (1895). Contributions to the mathematical theory of evolution. II. Skew variation in homogeneous material. <em>Philosophical Transactions of the Royal Society of London. A, </em>343-414.<br></br>" & _
        '"Richardson, M. W., & Stalnaker, J. M. (1933). A note on the use of bi-serial r in test research. <em>The Journal of General Psychology,</em> 8(2), 463-465.<br></br>" & _
        '"Yu, C. H., & Ds, P. (2012). A Simple Guide to the Item Response Theory (IRT) and Rasch Modeling.<br></br>Zeng, J., & Wyse, A. (2009). Introduction to Classical Test Theory. <em>Michigan, Washington, US.</em>" & _
        '"</td></tr>"
        strHTML &= "</tbody></table>"
        strHTML &= "</BODY></HTML>"

        System.IO.File.WriteAllText(oFileDialog.FileName.Substring(0, oFileDialog.FileName.IndexOf(".")) & ".htm", strHTML)
        Process.Start(oFileDialog.FileName.Substring(0, oFileDialog.FileName.IndexOf(".")) & ".htm")
    End Sub









    Private Sub cmdLoad_Click(sender As Object, e As EventArgs) Handles cmdLoad.Click
        If oFileDialog.FileName.Length > 0 Then
            dData.Clear()
            ReDim RawData(1, 1)
            ReDim BinData(1, 1)
            colSkip = -1
            numVar = 0
            numObs = 0
        End If
        oFileDialog.ShowDialog()
        If oFileDialog.FileName.Length > 0 Then
            lblStatus1.Text = "Loading file..."
            ReadDataFile(oFileDialog.FileName)
            lblStatus2.Text = "DONE"
        Else
            Exit Sub
        End If

        lblStatus1.Text = "Converting to Dichotomy..."

        CreateBinaryDataField()
        lblStatus2.Text = "DONE"
        lblStatus1.Text = "Calculating Item Difficulty..."
        Application.DoEvents()
        CalculatePValue()
        lblStatus2.Text = "DONE"
        lblStatus1.Text = "Calculating Standard Deviation..."
        Application.DoEvents()
        CalculateSTDev()
        lblStatus2.Text = "DONE"
        CalculateAlpha()
        lblStatus1.Text = "Calculating PBS Values..."
        Application.DoEvents()
        CalculatePointBiSerial()
        lblStatus2.Text = "DONE"
        lblStatus1.Text = "Calculating Descriptive Statistics..."
        Application.DoEvents()
        CalculateDescriptiveStats(stuRawScore)
        lblStatus2.Text = "DONE"
        lblStatus1.Text = "Calculating Test Reliability..."
        Application.DoEvents()
        CalculateAlphaIfDropped()
        lblStatus2.Text = "DONE"
        lblStatus1.Text = "Generating Report..."
        Application.DoEvents()
        HTMLOut()
        lblStatus2.Text = "DONE"
    End Sub


    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        End
    End Sub
End Class

