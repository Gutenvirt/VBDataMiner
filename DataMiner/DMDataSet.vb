'VBDataMiner - Extract and analyze data from MS Excel(c) files.
'Copyright (C) 2015 Chris Stefancik gutenvirt@gmail.com

'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'any later version.

'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.

'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.


Imports System.Data
Imports System.Data.DataSet
Imports System.IO

Public Class DMDataSet
    Const NullValue As String = "NaN"
    Dim sTime As New Stopwatch

    Public _ddata As New DataTable
    Dim BinData(,) As String
    Dim StuRawScoreBins(10) As Integer

    Public sTestName As String = ""
    Public sFileName As String

    Dim nNumberStudentsDrop As Integer = 0
    Dim nNumberColumns As Integer
    Dim nNumberRows As Integer

    Dim itemDifficulty() As Double
    Dim itemDiscrimination() As Double

    Dim c_CRPassRate() As Double

    Dim sCustomTitlePart As String = ""

    Dim ItemIntData(,) As Integer
    Dim ItemDblData(,) As Double
    Dim itemStrData(,) As String

    Dim stuRawScore() As Long

    Dim colSkip As Integer = -1

    Dim testMedian As Double
    Dim testQ1 As Double
    Dim testQ3 As Double

    Dim testSTDev As Double
    Dim testMean As Double
    Dim testAlpha As Double
    Dim testSTDEM As Double
    Dim testSkewness As Double

    Dim hasMC As Boolean = False
    Dim hasMS As Boolean = False
    Dim hasGR As Boolean = False
    Dim hasCR As Boolean = False


    Public Sub Initialize(ByVal DatabaseFile As String)
        sTime.Start()
        Const FIRST_ROW_LOC = 5
        Const FIRST_COL_LOC = 6
        Const LAST_COL_LOC = 5

        sFileName = DatabaseFile

        Dim _x As Integer = 0

        Try
            Dim _oleConnection As Data.OleDb.OleDbConnection = New Data.OleDb.OleDbConnection()
            Dim _oleAdapter As Data.OleDb.OleDbDataAdapter = New Data.OleDb.OleDbDataAdapter("SELECT * FROM [Sheet1$]", "provider=Microsoft.ACE.OLEDB.12.0; Data Source='" & DatabaseFile & "'; Extended Properties='Excel 12.0;IMEX=1;HDR=NO'")

            _oleAdapter.Fill(_ddata)
            _oleConnection.Close()

            _ddata.AcceptChanges()
        Catch e As Exception
            MsgBox("There was an unrecoverable error reading the file", vbCritical)
            Exit Sub
        End Try

        If _ddata.Rows(6).Item(1).ToString.IndexOf("1") < 0 Then
            MsgBox("This is not a correctly-formated file.", vbCritical)
            Exit Sub
        End If

        sTestName = _ddata.Rows(0).Item(6).ToString

        _x = 0
        While _x < FIRST_COL_LOC
            _ddata.Columns.RemoveAt(0)
            _x += 1
        End While

        _x = 0
        While _x < LAST_COL_LOC
            _ddata.Columns.RemoveAt(_ddata.Columns.Count - 1)
            _x += 1
        End While

        ReDim ItemIntData(_ddata.Columns.Count, 5)
        ReDim ItemDblData(_ddata.Columns.Count, 3)
        ReDim itemStrData(_ddata.Columns.Count, 3)

        _x = 0
        While _x < FIRST_ROW_LOC
            _ddata.Rows.RemoveAt(0)
            If _x = FIRST_ROW_LOC - 2 Then
                For _i = 0 To _ddata.Columns.Count - 1
                    itemStrData(_i, StrData.Standard) = _ddata.Rows(0).Item(_i).ToString
                Next
            End If
            _x += 1
        End While

        RemoveNullReferences()
        DetermineItemType()

        ConvertToDichotomy()
        GetMCFrequencies()
        CalculateAlphaIfDropped()
        CalculatePointBiSerial()
        CalculateDescriptiveStats(stuRawScore)

        _ddata.Dispose()
    End Sub

    Public Sub RemoveNullReferences()

        Dim isNullRow As Boolean = False
        For i As Integer = _ddata.Rows.Count - 1 To 0 Step -1
            isNullRow = False
            For j As Integer = _ddata.Columns.Count - 1 To 0 Step -1
                If _ddata.Rows(i).Item(j) Is Nothing Then
                    isNullRow = True
                Else
                    If IsDBNull(_ddata.Rows(i).Item(j)) Then
                        isNullRow = True
                    Else
                        If _ddata.Rows(i).Item(j).ToString = "" Then
                            isNullRow = True
                        Else
                            If _ddata.Rows(i).Item(j).ToString = "No" _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("MAFS.") > -1 _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("LAFS.") > -1 _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("SS.") > -1 _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("SC.") > -1 _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("%") > -1 _
                                Or _ddata.Rows(i).Item(j).ToString.IndexOf("Yes") > -1 _
                            Then
                                isNullRow = True
                                Exit For
                            Else
                                isNullRow = False
                                Exit For
                            End If
                        End If
                    End If
                End If
            Next
            If isNullRow = True Then
                nNumberStudentsDrop += 1
                _ddata.Rows.RemoveAt(i)
            End If
        Next

        nNumberColumns = _ddata.Columns.Count
        nNumberRows = _ddata.Rows.Count
        _ddata.AcceptChanges()
    End Sub

    Public Sub DetermineItemType()

        Dim _alpha As String = "ABCDFGHJ"
        Dim _TypeString As String = "MC"

        For i = 0 To nNumberColumns - 1
            _TypeString = "MC"
            For j = 0 To nNumberRows - 1
                If Not IsDBNull(_ddata.Rows(j).Item(i)) _
                    And Not IsNothing(_ddata.Rows(j).Item(i)) _
                    And _ddata.Rows(j).Item(i).ToString.IndexOf(",") > -1 Then
                    _TypeString = "MS"
                    hasMS = True
                    Exit For
                End If
            Next

            For j = 0 To nNumberRows - 1
                If _TypeString <> "MC" Then Exit For
                If Not IsDBNull(_ddata.Rows(j).Item(i)) _
                    And Not IsNothing(_ddata.Rows(j).Item(i)) _
                    And (_ddata.Rows(j).Item(i).ToString.IndexOf("-") > -1 Or _ddata.Rows(j).Item(i).ToString.IndexOf(".") > -1) Then
                    _TypeString = "GR"
                    hasGR = True
                    Exit For
                End If
            Next

            Dim NumPluses As Integer = 0
            For j = 0 To nNumberRows - 1
                If _TypeString <> "MC" Then Exit For
                If Not IsDBNull(_ddata.Rows(j).Item(i)) _
                    And Not IsNothing(_ddata.Rows(j).Item(i)) _
                    And _ddata.Rows(j).Item(i).ToString.IndexOf("+") > -1 _
                    And _ddata.Rows(j).Item(i).ToString <> "0" _
                    And IsNumeric(_ddata.Rows(j).Item(i).ToString) Then
                    NumPluses += 1
                End If
            Next

            If 1 - (NumPluses / nNumberRows) < 0.5 Then
                _TypeString = "CR"
                hasCR = True
            End If


            For j = 0 To nNumberRows - 1
                If _TypeString <> "MC" Then Exit For
                If Not IsDBNull(_ddata.Rows(j).Item(i)) _
                    And Not IsNothing(_ddata.Rows(j).Item(i)) _
                    And _ddata.Rows(j).Item(i).ToString.IndexOf("+") = -1 _
                    And IsNumeric(_ddata.Rows(j).Item(i).ToString) Then
                    _TypeString = "GR"
                    hasGR = True
                    Exit For
                End If
            Next

            If _TypeString = "MC" Then hasMC = True
            itemStrData(i, StrData.Type) = _TypeString
        Next

    End Sub

    Public Function ColumnAverage(ColumnIndex As Integer) As Double
        Dim _tAvg As Double = 0
        For _ix = 0 To nNumberRows - 1
            If Not _ddata.Rows(_ix).Item(ColumnIndex) Is Nothing _
                And IsDBNull(_ddata.Rows(_ix).Item(ColumnIndex)) = False Then
                _tAvg += CInt(_ddata.Rows(_ix).Item(ColumnIndex).ToString.Replace("+", ""))
            End If
        Next
        Return _tAvg / nNumberRows
    End Function

    Public Sub GetMCFrequencies()

        For i = 0 To nNumberColumns - 1
            For j = 0 To nNumberRows - 1
                If itemStrData(i, StrData.Type) <> "MC" Then Exit For

                If Not _ddata.Rows(j).Item(i) Is Nothing Then
                    Select Case _ddata.Rows(j).Item(i).ToString.Replace("+", "")
                        Case "A", "F"
                            ItemIntData(i, IntData.MC1) += 1
                        Case "B", "G"
                            ItemIntData(i, IntData.MC2) += 1
                        Case "C", "H"
                            ItemIntData(i, IntData.MC3) += 1
                        Case "D", "J"
                            ItemIntData(i, IntData.MC4) += 1
                    End Select
                End If
            Next
        Next

    End Sub

    Public Sub ConvertToDichotomy()


        ReDim BinData(nNumberColumns, nNumberRows)

        ReDim c_CRPassRate(nNumberColumns)

        Dim _t As String

        For i = 0 To nNumberColumns - 1
            If itemStrData(i, StrData.Type) = "CR" Then itemStrData(i, StrData.Answer) = ColumnAverage(i).ToString

            For j = 0 To nNumberRows - 1

                Try
                    If _ddata.Rows(j).Item(i) Is Nothing Or IsDBNull(_ddata.Rows(j).Item(i)) Then
                        BinData(i, j) = NullValue
                        ItemIntData(i, IntData.Omissions) += 1
                    Else
                        _t = _ddata.Rows(j).Item(i).ToString

                        If itemStrData(i, StrData.Type) = "CR" Then
                            If _t(0) <> "+" Then ItemIntData(i, IntData.Omissions) += 1
                            If CInt(_t.Replace("+", "")) >= CSng(itemStrData(i, StrData.Answer)) Then
                                BinData(i, j) = "1"
                                c_CRPassRate(i) += 1
                            Else
                                BinData(i, j) = "0"
                            End If
                        Else
                            If _t.IndexOf("+") > -1 Then
                                BinData(i, j) = "1"
                                If itemStrData(i, StrData.Answer) = "" Then itemStrData(i, StrData.Answer) = _t.Replace("+", "")
                            Else
                                BinData(i, j) = "0"
                                If itemStrData(i, StrData.Type) = "GR" Then
                                    If CInt(_t.Replace("+", "")) > CSng(itemStrData(i, StrData.Answer)) Then c_CRPassRate(i) += 1
                                End If
                            End If
                        End If
                    End If
                Catch
                    MsgBox("Error @ Col: " & i & ", Row: " & j & ", Type: " & itemStrData(i, StrData.Type))
                End Try
            Next
        Next
        For _a = 0 To nNumberColumns - 1
            c_CRPassRate(_a) = c_CRPassRate(_a) / nNumberRows * 100
        Next


    End Sub

    Public Sub CalculateSTDev()
        Dim _sumScoreVariance2 As Double
        ReDim stuRawScore(nNumberRows)

        For _a = 0 To nNumberRows - 1
            For _b = 0 To nNumberColumns - 1
                If BinData(_b, _a) = "1" And colSkip <> _b Then stuRawScore(_a) += CInt(BinData(_b, _a))
            Next
        Next

        testMean = stuRawScore.Average

        For _a = 0 To nNumberRows - 1
            _sumScoreVariance2 += Math.Pow(stuRawScore(_a) - testMean, 2)
        Next

        testSTDev = CSng(Math.Sqrt(_sumScoreVariance2 / nNumberRows))
    End Sub

    Public Sub CalculatePValue()

        Dim nNumberScorableItems(nNumberColumns) As Integer
        Dim _tPvCR(nNumberRows) As Integer

        For _a = 0 To nNumberColumns - 1
            For _b = 0 To nNumberRows - 1

                If BinData(_a, _b) <> NullValue And colSkip <> _a Then
                    If itemStrData(_a, StrData.Type) = "CR" And _ddata.Rows(_b).Item(_a).ToString.IndexOf("+") > -1 Then
                        nNumberScorableItems(_a) = 1
                        _tPvCR(_b) = CInt(_ddata.Rows(_b).Item(_a).ToString.Replace("+", ""))
                    Else
                        ItemDblData(_a, DblData.PV) += CInt(BinData(_a, _b))
                        nNumberScorableItems(_a) += 1
                    End If

                End If
            Next
            If itemStrData(_a, StrData.Type) = "CR" Then ItemDblData(_a, DblData.PV) = _tPvCR.Average / (_tPvCR.Max - _tPvCR.Min)

        Next

        ReDim itemDifficulty(3)
        For _a = 0 To nNumberColumns - 1
            If colSkip <> _a And itemStrData(_a, StrData.Type) <> "CR" Then ItemDblData(_a, DblData.PV) = ItemDblData(_a, DblData.PV) / nNumberScorableItems(_a)

            Select Case ItemDblData(_a, DblData.PV)
                Case Is <= 0.4
                    itemDifficulty(0) += 1
                Case 0.4 To 0.7
                    itemDifficulty(1) += 1
                Case 0.7
                    itemDifficulty(1) += 1
                Case Is > 0.7
                    itemDifficulty(2) += 1

            End Select
        Next

    End Sub

    Public Sub CalculateAlpha()
        Dim _sumPVariance As Double = 0

        For _a = 0 To nNumberColumns - 1
            If colSkip <> _a Then _sumPVariance += ItemDblData(_a, DblData.PV) * (1 - ItemDblData(_a, DblData.PV))
        Next
        testAlpha = CSng((1 / (nNumberColumns - 1) + 1) * (1 - _sumPVariance / Math.Pow(testSTDev, 2)))
        testSTDEM = CSng(testSTDev * Math.Sqrt(1 - testAlpha))
    End Sub

    Public Sub CalculatePointBiSerial()

        Dim _meanCorrect(nNumberColumns) As Double
        Dim _meanWrong(nNumberColumns) As Double

        Dim _numStuCorrect(nNumberColumns) As Double
        Dim _numStuWrong(nNumberColumns) As Double

        For _a = 0 To nNumberColumns - 1
            For _b = 0 To nNumberRows - 1
                Select Case BinData(_a, _b)
                    Case "1"
                        _meanCorrect(_a) += stuRawScore(_b)
                        _numStuCorrect(_a) += 1
                    Case Else
                        _meanWrong(_a) += stuRawScore(_b)
                        _numStuWrong(_a) += 1
                End Select
            Next
        Next

        ReDim itemDiscrimination(3)
        For _a = 0 To nNumberColumns - 1
            _meanCorrect(_a) = _meanCorrect(_a) / _numStuCorrect(_a)
            _meanWrong(_a) = _meanWrong(_a) / _numStuWrong(_a)
            ItemDblData(_a, DblData.PBS) = CSng((_meanCorrect(_a) - _meanWrong(_a)) / testSTDev * Math.Sqrt(_numStuCorrect(_a) / nNumberRows * _numStuWrong(_a) / nNumberRows))

            Select Case ItemDblData(_a, DblData.PBS)
                Case Is < 0.2
                    itemDiscrimination(0) += 1
                Case 0.2
                    itemDiscrimination(1) += 1
                Case 0.2 To 0.3
                    itemDiscrimination(1) += 1
                Case Is >= 0.3
                    itemDiscrimination(2) += 1
            End Select
        Next

    End Sub

    Public Sub CalculateAlphaIfDropped()

        For _h = 0 To nNumberColumns 'one more iteration with a non existant value of colsip = numcol +1

            colSkip = _h

            CalculateSTDev()
            CalculatePValue()
            CalculateAlpha()

            ItemDblData(_h, DblData.AIfD) = testAlpha

        Next

    End Sub

    Public Sub CalculateDescriptiveStats(StudentScores() As Long)
        Array.Sort(StudentScores)
        If StudentScores.Length Mod 2 <> 0 Then
            testMedian = StudentScores(StudentScores.GetUpperBound(0) \ 2)
            testQ1 = StudentScores(StudentScores.GetUpperBound(0) \ 4)
            testQ3 = StudentScores(3 * StudentScores.GetUpperBound(0) \ 4)
        Else
            testMedian = (StudentScores((StudentScores.Length \ 2)) + StudentScores((StudentScores.Length \ 2) - 1)) \ 2
            testQ1 = (StudentScores(StudentScores.Length \ 4) + StudentScores((StudentScores.Length \ 4) - 1)) \ 2
            testQ3 = (StudentScores(3 * StudentScores.Length \ 4) + StudentScores(3 * (StudentScores.Length \ 4) - 1)) \ 2
        End If

        testSkewness = 3 * (testMean - testMedian) / testSTDev

        Dim _s As Double

        For m = 0 To StudentScores.Length - 1

            _s = StudentScores(m) / nNumberColumns

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

    Public Function HTMLOutDistrict() As String
        Dim strHTML As String

        Dim tableWidth As Integer = 660 '488
        Dim gDivHeight As Integer = 250
        Dim gDivWidth As Integer = 515
        Dim barWidth As Integer = 45
        Dim medianLeft As Integer = 11 + CInt(testMean / nNumberColumns * gDivWidth)
        Dim medianHeight As Integer = 231
        Dim stdLeft As Integer = 11 + CInt((testMean - testSTDev) / nNumberColumns * gDivWidth)
        Dim stdWidth As Integer = CInt(testSTDev / nNumberColumns * 2 * gDivWidth)
        Dim stdTop As Integer = 234
        Dim p100Top As Integer = 0
        Dim p75Top As Integer = CInt(gDivHeight / 4)
        Dim p50Top As Integer = CInt(gDivHeight / 2)
        Dim p25Top As Integer = CInt(gDivHeight / 4 * 3)

        strHTML = "<!DOCTYPE html PUBLIC ""-//W3C//DTD HTML 4.01//EN""><HTML><HEAD><meta http-equiv=""Content-Type"" content=""text/html;charset=utf-8""><meta name=""author"" content=""Chris Stefancik 2015""><TITLE>" & sTestName & "</TITLE>"

        strHTML &= "<STYLE  type=""text/css"">"

        strHTML &= "table { border-collapse: collapse; border-color: #c1c1c1; border-spacing: 0; border-style: solid; border-width: 1px 0 0 1px; vertical-align: middle; width: " & tableWidth & "px; }"
        strHTML &= "th { background-color: #edf2f9; border-color: #b0b7bb; border-style: solid; border-width: 0 1px 1px 0; color: #112277; font-family: Arial, Helvetica, Helv; font-size: small; font-style: normal; font-weight: bold; padding: 3px 6px; text-align: center; vertical-align: middle; }"
        strHTML &= "td { background-color: #FFFFFF; border-color: #c1c1c1; border-style: solid; border-width: 0 1px 1px 0; font-family: Arial, Helvetica, Helv; font-size: small; font-style: normal; font-weight: normal; padding: 3px 6px; text-align: right; vertical-align: middle; }"
        strHTML &= ".graph { height: " & gDivHeight & "px; position: relative; width: " & gDivWidth & "px; }"
        strHTML &= ".bar { background-color: #edf2f9; border: 1px solid #c1c1c1; display: inline-block; margin: 1px; position: relative; vertical-align: baseline; width: " & barWidth & "px; }"
        strHTML &= ".median { background-color: #FBE2E0; border: 1px solid #9F9F9F; display: inline-block; height: " & medianHeight & "px; left: " & medianLeft & "px; margin: 0px; position: absolute; top: 1px; vertical-align: baseline; width: 0px; }"
        strHTML &= ".std { border: 1px solid #9F9F9F; display: inline-block; left: " & stdLeft & "px; margin: 0px; position: absolute; top: " & stdTop & "px; vertical-align: baseline; width: " & stdWidth & "px; }"
        strHTML &= ".xlabel { border: 1px solid #FFFFFF; display: inline-block; font-family: Arial, Helvetica, Helv; font-size: x-small; font-style: normal; font-weight: normal; margin: 1px; position: relative; text-align: center; vertical-align: baseline; width: " & barWidth & "px; }"
        strHTML &= ".ylabel { display: inline-block; font-family: Arial, Helvetica, Helv; font-size: x-small; font-style: normal; font-weight: normal; left: 0px; position: absolute; text-align: left; }"
        strHTML &= ".center { text-align: center; }"
        strHTML &= ".left { text-align: left; }"
        strHTML &= ".warning { background-color: #FBE2E0; }"

        strHTML &= "</STYLE>"

        strHTML &= "</HEAD><BODY><TABLE><tr><th colspan=""6""><p>" & sTestName.ToUpper & "</p></th><th colspan=""4""><p>Test Analysis Report</p></th></tr><tr><td class=""center"" colspan=""3"">Date: " & Date.Today & "</td><td class=""center"" colspan=""3"">User: " & Environment.UserName & "</td><td class=""center"" colspan=""4"">CDS 2015</td></tr>"

        strHTML &= "</table><p></p><table>"
        strHTML &= "<tr><th colspan=""2"">Raw Score</th><th colspan=""8"">Percent Score Distribution</th></tr>"

        If nNumberColumns < 3 Then
            strHTML &= "<tr><td>Items</td><td class=""warning"">" & nNumberColumns & "</td>"
        Else
            strHTML &= "<tr><td>Items</td><td>" & nNumberColumns & "</td>"
        End If

        strHTML &= "<td ROWSPAN=""12"" colspan=""8""><div><div class=""graph"">"
        For _a = 0 To 9
            strHTML &= "<div style=""height: " & CInt(StuRawScoreBins(_a) / StuRawScoreBins.Max * 230) & "px"" class=""bar""></div>"
        Next
        strHTML &= "<div class=""std""></div>"
        strHTML &= "<div class=""median""></div>"
        strHTML &= "<div class=""ylabel"" style=""top: " & p100Top & "px"">" & CInt(StuRawScoreBins.Max / nNumberRows * 100) & "%</div>"
        strHTML &= "<div class=""ylabel"" style=""top: " & p75Top & "px"">" & CInt(StuRawScoreBins.Max / nNumberRows * 75) & "%</div>"
        strHTML &= "<div class=""ylabel"" style=""top: " & p50Top & "px"">" & CInt(StuRawScoreBins.Max / nNumberRows * 50) & "%</div>"
        strHTML &= "<div class=""ylabel"" style=""top: " & p25Top & "px"">" & CInt(StuRawScoreBins.Max / nNumberRows * 25) & "%</div>"

        For _a = 0 To 9
            If _a = 9 Then
                strHTML &= "<div class=""xlabel"">" & _a * 10 & "-100</div>"
            Else
                strHTML &= "<div class=""xlabel"">" & _a * 10 & "-" & _a * 10 + 9 & "</div>"
            End If
        Next
        strHTML &= "</div></div></tr>"

        If nNumberRows < 25 Then
            strHTML &= "<tr><td>Students</td><td class=""warning"">" & nNumberRows & "</td></tr>"
        Else
            strHTML &= "<tr><td>Students</td><td>" & nNumberRows & "</td></tr>"
        End If

        strHTML &= "<tr><td>Min</td><td>" & stuRawScore.Min & "</td></tr>"
        strHTML &= "<tr><td>Q1</td><td>" & testQ1 & "</td></tr>"
        strHTML &= "<tr><td>Mean</td><td>" & testMean.ToString("0.00") & "</td></tr>"
        strHTML &= "<tr><td>Median</td><td>" & testMedian & "</td></tr>"
        strHTML &= "<tr><td>Q3</td><td>" & testQ3 & "</td></tr>"
        strHTML &= "<tr><td>Max</td><td>" & stuRawScore.Max & "</td></tr>"
        strHTML &= "<tr><td>Std Dev</td><td>" & testSTDev.ToString("0.00") & "</td></tr>"

        If testSkewness > 0 Then
            strHTML &= "<tr><td>Skew</td><td>&#8592; " & testSkewness.ToString("0.00") & "</td></tr>"
        Else
            strHTML &= "<tr><td>Skew</td><td>&#8594; " & testSkewness.ToString("0.00") & "</td></tr>"
        End If

        If testAlpha < 0.7 Or testAlpha > 1 Then
            strHTML &= "<tr><td>Alpha</td><td class=""warning"">" & testAlpha.ToString("0.00") & "</td></tr>"
        Else
            strHTML &= "<tr><td>Alpha</td><td>" & testAlpha.ToString("0.00") & "</td></tr>"
        End If

        strHTML &= "<tr><td>SEM</td><td>" & testSTDEM.ToString("0.00") & "</td></tr>"

        'Notes section
        If nNumberStudentsDrop > 0 Then
            strHTML &= "<tr><td class=""left"" colspan=""11"">"
            strHTML &= "Note: Some students (" & nNumberStudentsDrop & " total) were dropped because of complete test omission or an error in the data record; consider rescoring the test and re-running this analysis tool."
        End If
        strHTML &= "</table><p></p>"

        'Test Design Section

        strHTML &= "<table>"
        strHTML &= "<tr><th colspan=""4"">Item Difficulty</th>" & _
            "<th>% of Items</th>" & _
            "<th colspan=""4"">Item Discrimination</th>" & _
            "<th colspan=""1"">% of Items</th></tr>"

        strHTML &= "<tr><td colspan=""4"" class=""left"">Easy (Higher than 70%)</td>" & _
            "<td>" & Math.Round(itemDifficulty(2) / nNumberColumns * 100, 1) & "</td>" & _
            "<td colspan=""4"" class=""left"">Good (Higher than 0.3)</td>" & _
            "<td colspan=""1"">" & Math.Round(itemDiscrimination(2) / nNumberColumns * 100, 1) & "</td></tr>"

        strHTML &= "<tr><td colspan=""4"" class=""left"">Moderate (40% to 70%)</td>" & _
            "<td>" & Math.Round(itemDifficulty(1) / nNumberColumns * 100, 1) & "</td>" & _
            "<td colspan=""4"" class=""left"">Acceptable (0.2 to 0.3)</td>" & _
            "<td colspan=""1"">" & Math.Round(itemDiscrimination(1) / nNumberColumns * 100, 1) & "</td></tr>"

        strHTML &= "<tr><td colspan=""4"" class=""left"">Hard (Less than 40%)</td>" & _
            "<td>" & Math.Round(itemDifficulty(0) / nNumberColumns * 100, 1) & "</td>" & _
            "<td colspan=""4"" class=""left"">Needs Review (Less than 0.2)</td>" & _
            "<td colspan=""1"">" & Math.Round(itemDiscrimination(0) / nNumberColumns * 100, 1) & "</td></tr>"

        strHTML &= "</table><p></p>"

        'Item Review Section

        'Multiple Choice
        If hasMC = True Then
            strHTML &= "<table>"
            strHTML &= "<tr><th>Item</th><th>P-Value</th><th>PBS</th><th>Alpha IfD</th><th>Answer</th><th>% C1</th><th>% C2</th><th>% C3</th><th>% C4</th><th>% Om</th></tr>"


            For _a = 0 To nNumberColumns - 1
                If itemStrData(_a, StrData.Type) = "MC" Then
                    strHTML &= "<tr>"
                    strHTML &= "<td>" & _a + 1 & " " & itemStrData(_a, StrData.Type) & "</td>"
                    If ItemDblData(_a, DblData.PV) < 0.2 Or ItemDblData(_a, DblData.PV) > 0.9 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    End If
                    If ItemDblData(_a, DblData.PBS) < 0.2 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    End If

                    strHTML &= "<td>" & ItemDblData(_a, DblData.AIfD).ToString("0.00") & "</td>"
                    strHTML &= "<td class=""center"">" & itemStrData(_a, StrData.Answer) & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.MC1) / nNumberRows * 100, 0) & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.MC2) / nNumberRows * 100, 0) & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.MC3) / nNumberRows * 100, 0) & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.MC4) / nNumberRows * 100, 0) & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.Omissions) / nNumberRows * 100, 2) & "</td>"
                    strHTML &= "</tr>"
                End If
            Next
            strHTML &= "</table>"
            strHTML &= "<p></p>"
        End If


        'Multiple Select
        If hasMS = True Then
            strHTML &= "<table>"
            strHTML &= "<tr><th>Item</th><th>P-Value</th><th>PBS</th><th>Alpha IfD</th><th>Answer</th><th>% C1</th><th>% C2</th><th>% C3</th><th>% C4</th><th>% Om</th></tr>"

            For _a = 0 To nNumberColumns - 1
                If itemStrData(_a, StrData.Type) = "MS" Then
                    strHTML &= "<tr>"
                    strHTML &= "<td>" & _a + 1 & " " & itemStrData(_a, StrData.Type) & "</td>"
                    If ItemDblData(_a, DblData.PV) < 0.2 Or ItemDblData(_a, DblData.PV) > 0.9 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    End If
                    If ItemDblData(_a, DblData.PBS) < 0.2 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    End If

                    strHTML &= "<td>" & ItemDblData(_a, DblData.AIfD).ToString("0.00") & "</td>"
                    strHTML &= "<td class=""center"">" & itemStrData(_a, StrData.Answer) & "</td><td colspan=""4""></td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.Omissions) / nNumberRows * 100, 2) & "</td>"
                    strHTML &= "</tr>"
                End If
            Next
            strHTML &= "</table>"
            strHTML &= "<p></p>"
        End If

        'Gridded Response

        If hasGR = True Then
            strHTML &= "<table>"
            strHTML &= "<tr><th>Item</th><th>P-Value</th><th>PBS</th><th>Alpha IfD</th><th>Answer</th><th>Percent Below</th><th>Percent Above</th><th>% Om</th></tr>"

            For _a = 0 To nNumberColumns - 1
                If itemStrData(_a, StrData.Type) = "GR" Then
                    strHTML &= "<tr>"
                    strHTML &= "<td>" & _a + 1 & " " & itemStrData(_a, StrData.Type) & "</td>"
                    If ItemDblData(_a, DblData.PV) < 0.2 Or ItemDblData(_a, DblData.PV) > 0.9 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    End If
                    If ItemDblData(_a, DblData.PBS) < 0.2 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    End If

                    strHTML &= "<td>" & ItemDblData(_a, DblData.AIfD).ToString() & "</td>"
                    strHTML &= "<td class=""center"">" & itemStrData(_a, StrData.Answer) & "</td>"
                    strHTML &= "<td>" & (100 - ItemDblData(_a, DblData.PV) * 100 - c_CRPassRate(_a) - ItemIntData(_a, IntData.Omissions) / nNumberRows * 100).ToString("0.00") & "</td>"
                    strHTML &= "<td>" & c_CRPassRate(_a).ToString("0.00") & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.Omissions) / nNumberRows * 100, 2) & "</td>"
                    strHTML &= "</tr>"
                End If
            Next
            strHTML &= "</table>"
            strHTML &= "<p></p>"
        End If

        'Constructed Response
        If hasCR = True Then
            strHTML &= "<table>"
            strHTML &= "<tr><th>Item</th><th>P-Value</th><th>PBS</th><th>Alpha IfD</th><th>Mean</th><th>Percent Below</th><th>Percent at or Above</th><th>% Om</th></tr>"

            For _a = 0 To nNumberColumns - 1
                If itemStrData(_a, StrData.Type) = "CR" Then
                    strHTML &= "<tr>"
                    strHTML &= "<td>" & _a + 1 & " " & itemStrData(_a, StrData.Type) & "</td>"
                    If ItemDblData(_a, DblData.PV) < 0.2 Or ItemDblData(_a, DblData.PV) > 0.9 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PV).ToString("0.00") & "</td>"
                    End If
                    If ItemDblData(_a, DblData.PBS) < 0.2 Then
                        strHTML &= "<td class=""warning"">" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    Else
                        strHTML &= "<td>" & ItemDblData(_a, DblData.PBS).ToString("0.00") & "</td>"
                    End If

                    strHTML &= "<td>" & ItemDblData(_a, DblData.AIfD).ToString("0.00") & "</td>"
                    strHTML &= "<td class=""center"">" & itemStrData(_a, StrData.Answer).Substring(0, itemStrData(_a, StrData.Answer).IndexOf(".") + 3) & "</td>"
                    strHTML &= "<td>" & (100 - c_CRPassRate(_a) - (ItemIntData(_a, IntData.Omissions) / nNumberRows * 100)).ToString("0.00") & "</td>"
                    strHTML &= "<td>" & c_CRPassRate(_a).ToString("0.00") & "</td>"
                    strHTML &= "<td>" & Math.Round(ItemIntData(_a, IntData.Omissions) / nNumberRows * 100, 2) & "</td>"
                    strHTML &= "</tr>"
                End If
            Next
            strHTML &= "</table>"
            strHTML &= "<p></p>"
        End If

        'References Section

        If True Then
            strHTML &= "<table><tr><th colspan=""10"">Citations</th></tr><tr><td colspan=""10"" class=""left"">" & _
            "<p>Afifi, A. A., & Elashoff, R. M. (1966). Missing observations in multivariate statistics I. Review of the literature. <em>Journal of the American Statistical Association, </em> 61(315), 595-604.<br></br>" & _
            "Brown, J. D. (2001). Point-biserial correlation coefficients. <em>JALT Testing & Evaluation SIG Newsletter, </em> 5(3), 12-15.<br></br>" & _
            "Brown, S. (2011). Measures of shape: Skewness and Kurtosis. Retrieved on December, 31, 2014.<br></br>" & _
            "Ebel, R. L. (1950). Construction and validation of educational tests. <em>Review of Educational Research,</em> 87-97.<br></br>" & _
            "Ebel, R. L. (1965). Confidence Weighting and Test Reliability. <em>Journal of Educational Measurement,</em> 2(1), 49-57.<br></br>" & _
            "Kelley, T., Ebel, R., & Linacre, J. M. (2002). Item discrimination indices. <em>Rasch Measurement Transactions,</em> 16(3), 883-884.<br></br>" & _
            "Krishnan, V. (2013). The Early Child Development Instrument (EDI): An item analysis using Classical Test Theory (CTT) on Alberta’s data. <em>Early Child Mapping (ECMap) Project Alberta, Community-University Partnership (CUP), Faculty of Extension, University of Alberta, Edmonton, Alberta.</em><br></br>" & _
            "Matlock-Hetzel, S. (1997). Basic Concepts in Item and Test Analysis.<br></br>Pearson, K. (1895). Contributions to the mathematical theory of evolution. II. Skew variation in homogeneous material. <em>Philosophical Transactions of the Royal Society of London. A, </em>343-414.<br></br>" & _
            "Richardson, M. W., & Stalnaker, J. M. (1933). A note on the use of bi-serial r in test research. <em>The Journal of General Psychology,</em> 8(2), 463-465.<br></br>" & _
            "Yu, C. H., & Ds, P. (2012). A Simple Guide to the Item Response Theory (IRT) and Rasch Modeling.<br></br>Zeng, J., & Wyse, A. (2009). Introduction to Classical Test Theory. <em>Michigan, Washington, US.</em>" & _
            "</p></td></tr></table>"
        End If

        strHTML &= "</table><p></p>"

        sTime.Stop()
        strHTML &= "<table>"
        strHTML &= "<tr><th>Technical Information</th></td><tr><td class=""left"">Source file: " & sFileName.Remove(0, sFileName.LastIndexOf("\") + 1) & "<p></p>Total Elapsed Time: " & sTime.ElapsedMilliseconds / 1000 & " seconds</td></tr>"
        strHTML &= "</table></HTML>"

        Return strHTML
    End Function

End Class

Enum IntData
    Omissions
    MC1
    MC2
    MC3
    MC4
End Enum

Enum DblData
    PBS
    PV
    AIfD
End Enum

Enum StrData
    Type
    Answer
    Standard
End Enum