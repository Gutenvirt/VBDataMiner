﻿'VBDataMiner - Extract and analyze data from MS Excel(c) files.
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

Option Explicit On

Imports System.IO

Public Class frmMain
    Dim fdata As New DMDataSet
    Dim sName As String = ""

    Private Sub cmdLoad_Click(sender As Object, e As EventArgs) Handles cmdLoad.Click
        oFileDialog.ShowDialog()
        If oFileDialog.FileName.Length > 0 Then
            lblStatus1.Text = oFileDialog.FileName.Remove(0, oFileDialog.FileName.LastIndexOf("\") + 1)
            lblStatus2.Text = "Analyzing"
            Application.DoEvents()
            fdata.Initialize(oFileDialog.FileName)
            sName = fdata.sTestName.Replace(":", "-")
            Reporter(fdata.HTMLOutDistrict, oFileDialog.FileName)
            End
        End If
    End Sub

    Private Sub cmdExit_Click(sender As Object, e As EventArgs) Handles cmdExit.Click
        End
    End Sub

    Private Sub cmdLoadFolder_Click(sender As Object, e As EventArgs) Handles cmdLoadFolder.Click
        oFolderDialog.ShowDialog()
        If oFolderDialog.SelectedPath <> "" Then
            Dim di As New IO.DirectoryInfo(oFolderDialog.SelectedPath)
            Dim diar1 As IO.FileInfo() = di.GetFiles("*.xlsx")
            Dim dra As IO.FileInfo
            Dim _xCount As Integer = 1
            For Each dra In diar1
                Dim fData As New DMDataSet
                lblStatus1.Text = dra.Name
                lblStatus2.Text = _xCount & " of " & diar1.Count
                Application.DoEvents()
                fData.Initialize(dra.FullName.ToString)
                sName = fData.sTestName.Replace(":", "-")
                Reporter(fData.HTMLOutDistrict, dra.FullName)
                _xCount += 1
            Next
            End
        End If
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        cbAfterAnalysis.SelectedIndex = 0
    End Sub

    Public Sub Reporter(s_html As String, s_Filename As String)
        If s_Filename.IndexOf("ExportView") > -1 Then
            System.IO.File.WriteAllText(s_Filename.Substring(0, s_Filename.LastIndexOf("\") + 1) & sName.Replace(" ", "_") & ".htm", s_html)
            If cbAfterAnalysis.SelectedIndex = 0 Then Process.Start(s_Filename.Substring(0, s_Filename.LastIndexOf("\") + 1) & sName.Replace(" ", "_") & ".htm")
        Else
            System.IO.File.WriteAllText(s_Filename.Substring(0, s_Filename.IndexOf(".")) & ".htm", s_html)
            If cbAfterAnalysis.SelectedIndex = 0 Then Process.Start(s_Filename.Substring(0, s_Filename.IndexOf(".")) & ".htm")
        End If

    End Sub

End Class

