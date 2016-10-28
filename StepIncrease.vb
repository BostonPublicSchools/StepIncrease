Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Configuration
Imports System.Data
Imports System.Data.OleDb

Public Class StepIncrease

#Region "Private Properties"

    Private fileHistory As New OpenFileDialog
    Private fileUnion As New OpenFileDialog
    Private fileOutput As New SaveFileDialog

#End Region


#Region "Command Buttons Events"

    Private Sub cmdHistoryBrowse_Click(sender As System.Object, e As System.EventArgs) Handles cmdHistoryBrowse.Click

        fileHistory.Title = "Open History File"
        fileHistory.Filter = ConfigurationManager.AppSettings("defaultFilter")
        fileHistory.InitialDirectory = ConfigurationManager.AppSettings("defaultHistoryDirectory")


        If fileHistory.ShowDialog = Windows.Forms.DialogResult.OK Then
            lblHistory.Text = fileHistory.FileName.ToString()
        End If

    End Sub

    Private Sub cmdUnionBrowse_Click(sender As System.Object, e As System.EventArgs) Handles cmdUnionBrowse.Click

        fileUnion.Title = "Open Union File"
        fileUnion.Filter = ConfigurationManager.AppSettings("defaultFilter")
        fileUnion.InitialDirectory = ConfigurationManager.AppSettings("defaultUnionDirectory")

        If fileUnion.ShowDialog = Windows.Forms.DialogResult.OK Then
            lblUnion.Text = fileUnion.FileName.ToString()
        End If

    End Sub

    Private Sub cmdOutputBrowse_Click(sender As System.Object, e As System.EventArgs) Handles cmdOutputBrowse.Click

        fileOutput.Title = "Open Output File"
        fileOutput.Filter = ConfigurationManager.AppSettings("defaultFilter")
        fileOutput.InitialDirectory = ConfigurationManager.AppSettings("defaultOutputDirectory")

        If fileOutput.ShowDialog = Windows.Forms.DialogResult.OK Then
            lblOutput.Text = fileOutput.FileName.ToString()
        End If

    End Sub

    Private Sub cmdExit_Click(sender As System.Object, e As System.EventArgs) Handles cmdExit.Click

        Me.Close()

    End Sub

#End Region

#Region "Process File"

    Private Sub cmdProcessFiles_Click(sender As System.Object, e As System.EventArgs) Handles cmdProcessFiles.Click

        Dim cnn As OleDbConnection = New OleDbConnection(ConfigurationManager.AppSettings("oleDBConnectionString"))
        Dim cmd As OleDbCommand
        'Declare Variables
        Dim xlApp As Excel.Application
        Dim xlWB_Merged As Excel.Workbook
        Dim xlWS_Merged As Excel.Worksheet
        Dim xlWB_Output As Excel.Workbook
        Dim xlWS_Output As Excel.Worksheet
        Dim xlWS_Output_AFU As Excel.Worksheet
        Dim xlWS_Output_PAA As Excel.Worksheet
        Dim xlWS_Output_SU7 As Excel.Worksheet
        Dim xlWS_Output_SCA As Excel.Worksheet
        Dim xlWS_Output_BPS As Excel.Worksheet
        Dim xlWS_Output_SCATemp As Excel.Worksheet
        Dim EmplID As String
        Dim Found As Boolean
        Dim ServiceDate As Date
        Dim OutputLineAGU As Integer
        Dim OutputLineAFU As Integer
        Dim OutputLinePAA As Integer
        Dim OutputLineSU7 As Integer
        Dim OutputLineBPS As Integer
        Dim OutputLineSCA As Integer
        Dim OutputLineSCATemp As Integer
        Dim ProcessMonth As String = ""
        Dim x As Integer
        Dim OutputLine As Integer
        Dim MonthToProcess As Integer
        Dim nTotalRecords As Integer

        Windows.Forms.Cursor.Current = Cursors.WaitCursor

        For Each ctl As RadioButton In GroupBox1.Controls
            If ctl.Checked Then
                ProcessMonth = ctl.Text
                MonthToProcess = ctl.Tag
            End If
        Next

        If ProcessMonth = "" Then
            MsgBox("Please Select A Month To Process", MsgBoxStyle.OkOnly)   
            Exit Sub
        End If

        If System.IO.File.Exists(ConfigurationManager.AppSettings("defaultMergedDirectory") & "\Merged.xls") Then
            System.IO.File.Delete(ConfigurationManager.AppSettings("defaultMergedDirectory") & "\Merged.xls")
        End If

        cnn.Open()

        cmd = New OleDbCommand("SELECT * INTO [Union] FROM [Excel 8.0;DATABASE=" & fileUnion.FileName & "].[Sheet1$]", cnn)
        cmd.ExecuteNonQuery()
        cmd.CommandText = "SELECT * INTO [History] FROM [Excel 8.0;DATABASE=" & fileHistory.FileName & "].[Sheet1$]"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "SELECT DISTINCT [Union].[Pay Status], [Union].[Empl Class], [Union].[ID], [Union].[Name], [Union].[NID], [Union].[Grade], [Union].[Step], [Union].[Job Code], [Union].[Job Title], [Union].[DeptID], [Union].[Dept Title], [Union].[Union Code], [Union].[Union Date], [Union].[Prof Exp Date (Longevity)], [History].[Adj Pay Date], [History].[Init Perm Dt], [History].[Curr Perm Dt], [History].[Cur Perm Job], [Union].[Service Dt] INTO [Merged] FROM [Union] LEFT JOIN [History] ON [Union].[ID] = [History].[ID]"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "SELECT * INTO [Excel 8.0;Database=" & ConfigurationManager.AppSettings("defaultMergedDirectory") & "\Merged.xls].[Sheet1] FROM [Merged]"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Drop TABLE [Union]"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Drop TABLE [History]"
        cmd.ExecuteNonQuery()
        cmd.CommandText = "Select count(*) from [Merged]"
        nTotalRecords = cmd.ExecuteScalar() 
        cmd.CommandText = "Drop TABLE [Merged]"
        cmd.ExecuteNonQuery()
        cnn.Close()
        'Open Excel Files
        xlApp = New Excel.Application
        xlWB_Merged = xlApp.Workbooks.Open(ConfigurationManager.AppSettings("defaultMergedDirectory") & "\Merged.xls")
        xlWS_Merged = xlApp.Worksheets("Sheet1")
        Dim tempRange As Excel.Range
        tempRange = xlWS_Merged.Range("A1", "S65536")
        tempRange.Columns.AutoFit()
        xlWB_Merged.Save()
        xlWB_Output = xlApp.Workbooks.Add
        xlWS_Output = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "AGU"
        xlApp.Worksheets.Add()
        xlWS_Output_AFU = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "AFU"
         xlApp.Worksheets.Add()
        xlWS_Output_PAA = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "PAA"
        xlApp.Worksheets.Add()
        xlWS_Output_SU7 = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "SU7"
        xlApp.Worksheets.Add()
        xlWS_Output_SCA = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "SCA"
        xlApp.Worksheets.Add()
        xlWS_Output_BPS = xlApp.Application.ActiveSheet
        xlApp.Application.ActiveSheet.Name = "BPS"

        'Add Column Headings to Merged File
        x = 1
        xlWS_Merged.Cells(x, 20) = "Service Date"
        xlWS_Merged.Cells(x, 21) = "Date Source"

        'Add Column headings to Output File
        x = 1
        xlWS_Output.Cells(x, 1) = "ID"
        xlWS_Output.Cells(x, 2) = "Name"
        xlWS_Output.Cells(x, 3) = "SSN"
        xlWS_Output.Cells(x, 4) = "Grade"
        xlWS_Output.Cells(x, 5) = "Step"
        xlWS_Output.Cells(x, 6) = "Award Date"
        xlWS_Output.Cells(x, 7) = "Data Source"
        xlWS_Output.Cells(x, 8) = "Dept"
        xlWS_Output.Cells(x, 9) = "Title"
        xlWS_Output.Cells(x, 10) = "Award Type"

        xlWS_Output_AFU.Cells(x, 1) = "ID"
        xlWS_Output_AFU.Cells(x, 2) = "Name"
        xlWS_Output_AFU.Cells(x, 3) = "SSN"
        xlWS_Output_AFU.Cells(x, 4) = "Grade"
        xlWS_Output_AFU.Cells(x, 5) = "Step"
        xlWS_Output_AFU.Cells(x, 6) = "Award Date"
        xlWS_Output_AFU.Cells(x, 7) = "Data Source"
        xlWS_Output_AFU.Cells(x, 8) = "Dept"
        xlWS_Output_AFU.Cells(x, 9) = "Title"
        xlWS_Output_AFU.Cells(x, 10) = "Award Type"

        xlWS_Output_PAA.Cells(x, 1) = "ID"
        xlWS_Output_PAA.Cells(x, 2) = "Name"
        xlWS_Output_PAA.Cells(x, 3) = "SSN"
        xlWS_Output_PAA.Cells(x, 4) = "Grade"
        xlWS_Output_PAA.Cells(x, 5) = "Step"
        xlWS_Output_PAA.Cells(x, 6) = "Award Date"
        xlWS_Output_PAA.Cells(x, 7) = "Data Source"
        xlWS_Output_PAA.Cells(x, 8) = "Dept"
        xlWS_Output_PAA.Cells(x, 9) = "Title"
        xlWS_Output_PAA.Cells(x, 10) = "Award Type"

        xlWS_Output_SU7.Cells(x, 1) = "ID"
        xlWS_Output_SU7.Cells(x, 2) = "Name"
        xlWS_Output_SU7.Cells(x, 3) = "SSN"
        xlWS_Output_SU7.Cells(x, 4) = "Grade"
        xlWS_Output_SU7.Cells(x, 5) = "Step"
        xlWS_Output_SU7.Cells(x, 6) = "Award Date"
        xlWS_Output_SU7.Cells(x, 7) = "Data Source"
        xlWS_Output_SU7.Cells(x, 8) = "Dept"
        xlWS_Output_SU7.Cells(x, 9) = "Title"
        xlWS_Output_SU7.Cells(x, 10) = "Award Type"

        xlWS_Output_SCA.Cells(x, 1) = "ID"
        xlWS_Output_SCA.Cells(x, 2) = "Name"
        xlWS_Output_SCA.Cells(x, 3) = "SSN"
        xlWS_Output_SCA.Cells(x, 4) = "Grade"
        xlWS_Output_SCA.Cells(x, 5) = "Step"
        xlWS_Output_SCA.Cells(x, 6) = "Award Date"
        xlWS_Output_SCA.Cells(x, 7) = "Data Source"
        xlWS_Output_SCA.Cells(x, 8) = "Dept"
        xlWS_Output_SCA.Cells(x, 9) = "Title"
        xlWS_Output_SCA.Cells(x, 10) = "Award Type"
        xlWS_Output_SCA.Cells(x, 11) = "Curr Perm Date"

        xlWS_Output_BPS.Cells(x, 1) = "ID"
        xlWS_Output_BPS.Cells(x, 2) = "Name"
        xlWS_Output_BPS.Cells(x, 3) = "SSN"
        xlWS_Output_BPS.Cells(x, 4) = "Grade"
        xlWS_Output_BPS.Cells(x, 5) = "Step"
        xlWS_Output_BPS.Cells(x, 6) = "Professional Date"
        xlWS_Output_BPS.Cells(x, 7) = "Dept"
        xlWS_Output_BPS.Cells(x, 8) = "Title"

        OutputLine = 2


        'Set Variables
        x = 2
        ProgressBar1.Minimum = 2 
        ProgressBar1.Maximum = nTotalRecords + X
        OutputLineAGU = 2
        OutputLineAFU = 2
        OutputLinePAA = 2
        OutputLineSU7 = 2
        OutputLineBPS = 2
        OutputLineSCA = 2

        Do While (xlWS_Merged.Cells(x, 3).Text <> "")
            If (xlWS_Merged.Cells(x, 15).Text <> "") Then
                xlWS_Merged.Cells(x, 20).Value = CDate(xlWS_Merged.Cells(x, 15).Text)
                xlWS_Merged.Cells(x, 21).Value = "A"
            ElseIf (xlWS_Merged.Cells(x, 16).Text <> "") Then
                xlWS_Merged.Cells(x, 20).Value = CDate(xlWS_Merged.Cells(x, 16).Text)
                xlWS_Merged.Cells(x, 21).Value = "I"
            ElseIf (xlWS_Merged.Cells(x, 17).Text <> "") Then
                xlWS_Merged.Cells(x, 20).Value = CDate(xlWS_Merged.Cells(x, 17).Text)
                xlWS_Merged.Cells(x, 21).Value = "C"
            ElseIf (xlWS_Merged.Cells(x, 2).Text = "Q") And (xlWS_Merged.Cells(x, 12).Text = "SCA") And (xlWS_Merged.Cells(x, 19).Text <> "") Then
                xlWS_Merged.Cells(x, 20).Value = CDate(xlWS_Merged.Cells(x, 19).Text)
                xlWS_Merged.Cells(x, 21).Value = "B"
            Else
                '               Stop ' what do we put here in date field 20 if  none of above conditions are true?
            End If




            If ((xlWS_Merged.Cells(x, 12).Text = "BPS") And (MonthToProcess = GetMonth(xlWS_Merged.Cells(x, 14)))) Then

                xlWS_Output_BPS.Cells(OutputLineBPS, 1) = xlWS_Merged.Cells(x, 3)
                xlWS_Output_BPS.Cells(OutputLineBPS, 2) = xlWS_Merged.Cells(x, 4)
                xlWS_Output_BPS.Cells(OutputLineBPS, 3) = xlWS_Merged.Cells(x, 5)
                xlWS_Output_BPS.Cells(OutputLineBPS, 4) = xlWS_Merged.Cells(x, 6)
                xlWS_Output_BPS.Cells(OutputLineBPS, 5) = xlWS_Merged.Cells(x, 7)
                xlWS_Output_BPS.Cells(OutputLineBPS, 6) = xlWS_Merged.Cells(x, 14)
                xlWS_Output_BPS.Cells(OutputLineBPS, 7) = xlWS_Merged.Cells(x, 10)
                xlWS_Output_BPS.Cells(OutputLineBPS, 8) = xlWS_Merged.Cells(x, 11)
                OutputLineBPS = OutputLineBPS + 1

            ElseIf (MonthToProcess = GetMonth(xlWS_Merged.Cells(x, 20))) Then

                Select Case (xlWS_Merged.Cells(x, 12).Text)

                    Case "AGU"
                        'Check to see if person should get career award
                        Select Case (Year(ConfigurationManager.AppSettings("defaultDate")) - GetYear(xlWS_Merged.Cells(x, 20)))
                            Case "9", "14", "19", "24", "29", "34", "39"
                                xlWS_Output.Cells(OutputLineAGU, 1) = xlWS_Merged.Cells(x, 3)
                                xlWS_Output.Cells(OutputLineAGU, 2) = xlWS_Merged.Cells(x, 4)
                                xlWS_Output.Cells(OutputLineAGU, 3) = xlWS_Merged.Cells(x, 5)
                                xlWS_Output.Cells(OutputLineAGU, 4) = xlWS_Merged.Cells(x, 6)
                                xlWS_Output.Cells(OutputLineAGU, 5) = xlWS_Merged.Cells(x, 7)
                                xlWS_Output.Cells(OutputLineAGU, 6) = xlWS_Merged.Cells(x, 20)
                                xlWS_Output.Cells(OutputLineAGU, 7) = xlWS_Merged.Cells(x, 21)
                                xlWS_Output.Cells(OutputLineAGU, 8) = xlWS_Merged.Cells(x, 10)
                                xlWS_Output.Cells(OutputLineAGU, 9) = xlWS_Merged.Cells(x, 11)
                                xlWS_Output.Cells(OutputLineAGU, 10) = "Career Award"
                                OutputLineAGU = OutputLineAGU + 1
                        End Select
                        'Check to see if person should get step increase
                        If (xlWS_Merged.Cells(x, 7).Value < 7) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            xlWS_Output.Cells(OutputLineAGU, 1) = xlWS_Merged.Cells(x, 3)
                            xlWS_Output.Cells(OutputLineAGU, 2) = xlWS_Merged.Cells(x, 4)
                            xlWS_Output.Cells(OutputLineAGU, 3) = xlWS_Merged.Cells(x, 5)
                            xlWS_Output.Cells(OutputLineAGU, 4) = xlWS_Merged.Cells(x, 6)
                            xlWS_Output.Cells(OutputLineAGU, 5) = xlWS_Merged.Cells(x, 7)
                            xlWS_Output.Cells(OutputLineAGU, 6) = xlWS_Merged.Cells(x, 20)
                            xlWS_Output.Cells(OutputLineAGU, 7) = xlWS_Merged.Cells(x, 21)
                            xlWS_Output.Cells(OutputLineAGU, 8) = xlWS_Merged.Cells(x, 10)
                            xlWS_Output.Cells(OutputLineAGU, 9) = xlWS_Merged.Cells(x, 11)
                            xlWS_Output.Cells(OutputLineAGU, 10) = "Step Increase"
                            OutputLineAGU = OutputLineAGU + 1
                        End If

                    Case "AFU"
                        'Check to see if person should get career award
                        Select Case (Year(ConfigurationManager.AppSettings("defaultDate")) - GetYear(xlWS_Merged.Cells(x, 20)))
                            Case "9", "14", "19", "29", "39"
                                xlWS_Output_AFU.Cells(OutputLineAFU, 1) = xlWS_Merged.Cells(x, 3)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 2) = xlWS_Merged.Cells(x, 4)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 3) = xlWS_Merged.Cells(x, 5)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 4) = xlWS_Merged.Cells(x, 6)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 5) = xlWS_Merged.Cells(x, 7)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 6) = xlWS_Merged.Cells(x, 20)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 7) = xlWS_Merged.Cells(x, 21)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 8) = xlWS_Merged.Cells(x, 10)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 9) = xlWS_Merged.Cells(x, 11)
                                xlWS_Output_AFU.Cells(OutputLineAFU, 10) = "Career Award"
                                OutputLineAFU = OutputLineAFU + 1
                        End Select
                        'Check to see if person should get step increase
                        If (xlWS_Merged.Cells(x, 7).Value < 7) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            xlWS_Output_AFU.Cells(OutputLineAFU, 1) = xlWS_Merged.Cells(x, 3)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 2) = xlWS_Merged.Cells(x, 4)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 3) = xlWS_Merged.Cells(x, 5)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 4) = xlWS_Merged.Cells(x, 6)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 5) = xlWS_Merged.Cells(x, 7)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 6) = xlWS_Merged.Cells(x, 20)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 7) = xlWS_Merged.Cells(x, 21)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 8) = xlWS_Merged.Cells(x, 10)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 9) = xlWS_Merged.Cells(x, 11)
                            xlWS_Output_AFU.Cells(OutputLineAFU, 10) = "Step Increase"
                            OutputLineAFU = OutputLineAFU + 1
                        End If

                    Case "PAA"
                        'Check to see if person should get career award
                        Select Case (Year(ConfigurationManager.AppSettings("defaultDate")) - GetYear(xlWS_Merged.Cells(x, 20)))
                            Case "9", "14", "19", "29", "39"
                                xlWS_Output_PAA.Cells(OutputLinePAA, 1) = xlWS_Merged.Cells(x, 3)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 2) = xlWS_Merged.Cells(x, 4)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 3) = xlWS_Merged.Cells(x, 5)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 4) = xlWS_Merged.Cells(x, 6)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 5) = xlWS_Merged.Cells(x, 7)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 6) = xlWS_Merged.Cells(x, 20)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 7) = xlWS_Merged.Cells(x, 21)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 8) = xlWS_Merged.Cells(x, 10)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 9) = xlWS_Merged.Cells(x, 11)
                                xlWS_Output_PAA.Cells(OutputLinePAA, 10) = "Career Award"
                                OutputLinePAA = OutputLinePAA + 1
                        End Select
                        'Check to see if person should get step increase
                        If (xlWS_Merged.Cells(x, 7).Value < 5) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            xlWS_Output_PAA.Cells(OutputLinePAA, 1) = xlWS_Merged.Cells(x, 3)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 2) = xlWS_Merged.Cells(x, 4)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 3) = xlWS_Merged.Cells(x, 5)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 4) = xlWS_Merged.Cells(x, 6)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 5) = xlWS_Merged.Cells(x, 7)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 6) = xlWS_Merged.Cells(x, 20)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 7) = xlWS_Merged.Cells(x, 21)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 8) = xlWS_Merged.Cells(x, 10)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 9) = xlWS_Merged.Cells(x, 11)
                            xlWS_Output_PAA.Cells(OutputLinePAA, 10) = "Step Increase"
                            OutputLinePAA = OutputLinePAA + 1
                        End If

                    Case "SU7"
                        'Check to see if person should get career award
                        Select Case (Year(ConfigurationManager.AppSettings("defaultDate")) - GetYear(xlWS_Merged.Cells(x, 20)))
                            Case "9", "14", "19", "29", "39"
                                xlWS_Output_SU7.Cells(OutputLineSU7, 1) = xlWS_Merged.Cells(x, 3)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 2) = xlWS_Merged.Cells(x, 4)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 3) = xlWS_Merged.Cells(x, 5)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 4) = xlWS_Merged.Cells(x, 6)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 5) = xlWS_Merged.Cells(x, 7)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 6) = xlWS_Merged.Cells(x, 20)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 7) = xlWS_Merged.Cells(x, 21)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 8) = xlWS_Merged.Cells(x, 10)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 9) = xlWS_Merged.Cells(x, 11)
                                xlWS_Output_SU7.Cells(OutputLineSU7, 10) = "Career Award"
                                OutputLineSU7 = OutputLineSU7 + 1
                        End Select
                        'Check to see if person should get step increase
                        If (xlWS_Merged.Cells(x, 7).Value < 5) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            xlWS_Output_SU7.Cells(OutputLineSU7, 1) = xlWS_Merged.Cells(x, 3)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 2) = xlWS_Merged.Cells(x, 4)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 3) = xlWS_Merged.Cells(x, 5)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 4) = xlWS_Merged.Cells(x, 6)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 5) = xlWS_Merged.Cells(x, 7)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 6) = xlWS_Merged.Cells(x, 20)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 7) = xlWS_Merged.Cells(x, 21)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 8) = xlWS_Merged.Cells(x, 10)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 9) = xlWS_Merged.Cells(x, 11)
                            xlWS_Output_SU7.Cells(OutputLineSU7, 10) = "Step Increase"
                            OutputLineSU7 = OutputLineSU7 + 1
                        End If

                    Case "SCA"
                        'Career Awards (Jr's and Sr's)
                        Select Case (Year(ConfigurationManager.AppSettings("defaultDate")) - GetYear(xlWS_Merged.Cells(x, 20)))
                            Case "9", "14", "19", "24", "29", "34", "39"
                                xlWS_Output_SCA.Cells(OutputLineSCA, 1) = xlWS_Merged.Cells(x, 3)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 2) = xlWS_Merged.Cells(x, 4)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 3) = xlWS_Merged.Cells(x, 5)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 4) = xlWS_Merged.Cells(x, 6)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 5) = xlWS_Merged.Cells(x, 7)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 6) = xlWS_Merged.Cells(x, 20)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 7) = xlWS_Merged.Cells(x, 21)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 8) = xlWS_Merged.Cells(x, 10)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 9) = xlWS_Merged.Cells(x, 11)
                                xlWS_Output_SCA.Cells(OutputLineSCA, 10) = "Career Award"
                                xlWS_Output_SCA.Cells(OutputLineSCA, 11) = xlWS_Merged.Cells(x, 17)
                                OutputLineSCA = OutputLineSCA + 1
                        End Select
                        'Step Increase for Juniors
                        If (xlWS_Merged.Cells(x, 2).Text <> "Q") And (xlWS_Merged.Cells(x, 6).Text = "A") And (xlWS_Merged.Cells(x, 7).Value < 8) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            xlWS_Output_SCA.Cells(OutputLineSCA, 1) = xlWS_Merged.Cells(x, 3)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 2) = xlWS_Merged.Cells(x, 4)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 3) = xlWS_Merged.Cells(x, 5)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 4) = xlWS_Merged.Cells(x, 6)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 5) = xlWS_Merged.Cells(x, 7)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 6) = xlWS_Merged.Cells(x, 20)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 7) = xlWS_Merged.Cells(x, 21)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 8) = xlWS_Merged.Cells(x, 10)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 9) = xlWS_Merged.Cells(x, 11)
                            xlWS_Output_SCA.Cells(OutputLineSCA, 10) = "Jr Step Increase"
                            xlWS_Output_SCA.Cells(OutputLineSCA, 11) = xlWS_Merged.Cells(x, 17)
                            OutputLineSCA = OutputLineSCA + 1
                        End If

                        'Step Increase for Temp
                        If (xlWS_Merged.Cells(x, 2).Text = "Q") And (xlWS_Merged.Cells(x, 6).Text = "A") And (xlWS_Merged.Cells(x, 7).Value < 8) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                            Select Case (xlWS_Merged.Cells(x, 3).Value)
                                Case "077649", "091330", "098335", "096357", "101730", "099330", "100967", "096121", "092728", _
                                 "076155", "094428", "091682", "077120", "091714", "104424", "101733", "096299", "102513", "105697", _
                                 "099204", "104156", "099874", "088991", "103501", "049813", "099676", "095415", "096661", "096247", _
                                 "093938", "104629", "096246", "105600", "102732", "102961", "092905", "091868", "083120", "103094", _
                                 "096373", "103093", "089743", "077754", "089605", "092740", "099748", "106463", "103500", "051564", _
                                 "045330", "099873", "083508", "099553", "090914", "092075", "099188", "091863", "091653", "092560", _
                                 "097861", "092021", "091030", "096245", "095264", "091234", "090919", "098549", "094201", "074802", _
                                 "103107", "094693", "092501", "098962", "090924", "098423", "096971", "091759", "100995", "085752", _
                                 "098696", "100142", "094014", "097936", "096707", "102987", "091311", "081892", "091442", "103954", "077714", "096439", "055937"

                                    'DO Nothing

                                Case Else
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 1) = xlWS_Merged.Cells(x, 3)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 2) = xlWS_Merged.Cells(x, 4)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 3) = xlWS_Merged.Cells(x, 5)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 4) = xlWS_Merged.Cells(x, 6)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 5) = xlWS_Merged.Cells(x, 7)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 6) = xlWS_Merged.Cells(x, 20)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 7) = xlWS_Merged.Cells(x, 21)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 8) = xlWS_Merged.Cells(x, 10)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 9) = xlWS_Merged.Cells(x, 11)
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 10) = "Temp Step Increase"
                                    xlWS_Output_SCA.Cells(OutputLineSCA, 11) = xlWS_Merged.Cells(x, 17)
                                    OutputLineSCA = OutputLineSCA + 1
                            End Select
                        End If
                End Select
            End If

            'Step Increase for SCA Seniors
            If ((MonthToProcess = GetMonth(xlWS_Merged.Cells(x, 17))) And (GetYear(xlWS_Merged.Cells(x, 17)) < Year(ConfigurationManager.AppSettings("defaultDate"))) And (xlWS_Merged.Cells(x, 12).Text = "SCA") And (xlWS_Merged.Cells(x, 6).Text = "C") And (xlWS_Merged.Cells(x, 7).Value < 8) And ((xlWS_Merged.Cells(x, 18).Text = "407") Or (xlWS_Merged.Cells(x, 18).Text = "S40015") Or (xlWS_Merged.Cells(x, 18).Text = "S40040") Or (xlWS_Merged.Cells(x, 18).Text = "S40041") Or (xlWS_Merged.Cells(x, 18).Text = "S40050") Or (xlWS_Merged.Cells(x, 18).Text = "S40051") Or (xlWS_Merged.Cells(x, 18).Text = "S40060") Or (xlWS_Merged.Cells(x, 18).Text = "S40061"))) Then
                xlWS_Output_SCA.Cells(OutputLineSCA, 1) = xlWS_Merged.Cells(x, 3)
                xlWS_Output_SCA.Cells(OutputLineSCA, 2) = xlWS_Merged.Cells(x, 4)
                xlWS_Output_SCA.Cells(OutputLineSCA, 3) = xlWS_Merged.Cells(x, 5)
                xlWS_Output_SCA.Cells(OutputLineSCA, 4) = xlWS_Merged.Cells(x, 6)
                xlWS_Output_SCA.Cells(OutputLineSCA, 5) = xlWS_Merged.Cells(x, 7)
                xlWS_Output_SCA.Cells(OutputLineSCA, 6) = xlWS_Merged.Cells(x, 17)
                xlWS_Output_SCA.Cells(OutputLineSCA, 7) = "C"
                xlWS_Output_SCA.Cells(OutputLineSCA, 8) = xlWS_Merged.Cells(x, 10)
                xlWS_Output_SCA.Cells(OutputLineSCA, 9) = xlWS_Merged.Cells(x, 11)
                xlWS_Output_SCA.Cells(OutputLineSCA, 10) = "Sr Step Increase"
                xlWS_Output_SCA.Cells(OutputLineSCA, 11) = xlWS_Merged.Cells(x, 17)
                OutputLineSCA = OutputLineSCA + 1
            ElseIf (MonthToProcess = "7") And (xlWS_Merged.Cells(x, 2).Text = "Q") And (xlWS_Merged.Cells(x, 6).Text = "A") And (xlWS_Merged.Cells(x, 7).Value < 8) And (GetYear(xlWS_Merged.Cells(x, 20)) < Year(ConfigurationManager.AppSettings("defaultDate"))) Then
                Select Case (xlWS_Merged.Cells(x, 3).Value)
                    Case "077649", "091330", "098335", "096357", "101730", "099330", "100967", "096121", "092728", _
                         "076155", "094428", "091682", "077120", "091714", "104424", "101733", "096299", "102513", "105697", _
                         "099204", "104156", "099874", "088991", "103501", "049813", "099676", "095415", "096661", "096247", _
                         "093938", "104629", "096246", "105600", "102732", "102961", "092905", "091868", "083120", "103094", _
                         "096373", "103093", "089743", "077754", "089605", "092740", "099748", "106463", "103500", "051564", _
                         "045330", "099873", "083508", "099553", "090914", "092075", "099188", "091863", "091653", "092560", _
                         "097861", "092021", "091030", "096245", "095264", "091234", "090919", "098549", "094201", "074802", _
                         "103107", "094693", "092501", "098962", "090924", "098423", "096971", "091759", "100995", "085752", _
                         "098696", "100142", "094014", "097936", "096707", "102987", "091311", "081892", "091442", "103954", "077714", "096439", "055937"


                        xlWS_Output_SCA.Cells(OutputLineSCA, 1) = xlWS_Merged.Cells(x, 3)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 2) = xlWS_Merged.Cells(x, 4)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 3) = xlWS_Merged.Cells(x, 5)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 4) = xlWS_Merged.Cells(x, 6)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 5) = xlWS_Merged.Cells(x, 7)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 6) = xlWS_Merged.Cells(x, 20)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 7) = xlWS_Merged.Cells(x, 21)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 8) = xlWS_Merged.Cells(x, 10)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 9) = xlWS_Merged.Cells(x, 11)
                        xlWS_Output_SCA.Cells(OutputLineSCA, 10) = "Temp Step (July 1) Increase"
                        xlWS_Output_SCA.Cells(OutputLineSCA, 11) = xlWS_Merged.Cells(x, 17)
                        OutputLineSCA = OutputLineSCA + 1
                End Select
            End If

            x = x + 1
            ProgressBar1.Value = x
            Me.Text = "Processing Record " & x
            'Dim percent As Integer = CInt(Math.Truncate((CDbl(progressBar1.Value - progressBar1.Minimum) / CDbl(progressBar1.Maximum - progressBar1.Minimum)) * 100))
            'Using gr As Graphics = ProgressBar1.CreateGraphics()
	           ' gr.DrawString(percent.ToString() & "%", SystemFonts.DefaultFont, Brushes.Black, New PointF(progressBar1.Width / 2 - (gr.MeasureString(percent.ToString() & "%", SystemFonts.DefaultFont).Width / 2F), progressBar1.Height / 2 - (gr.MeasureString(percent.ToString() & "%", SystemFonts.DefaultFont).Height / 2F)))
            'End Using




        Loop

        'Format Worksheets
        Dim FormatString As String

        FormatString = "Administrative Guild Employees Due a Career Award or Step Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output, FormatString, 2)
        xlWS_Output.Range("F:F").NumberFormat = "m/d/yyyy;@"

        FormatString = "Storekeepers Due a Career Award or Step Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output_AFU, FormatString, 2)
        xlWS_Output_AFU.Range("F:F").NumberFormat = "m/d/yyyy;@"

        FormatString = "Plant Administrators Due a Career Award or Step Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output_PAA, FormatString, 2)
        xlWS_Output_PAA.Range("F:F").NumberFormat = "m/d/yyyy;@"

        FormatString = "Planning and Engineering Employees Due a Career Award or Step Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output_SU7, FormatString, 2)
        xlWS_Output_SU7.Range("F:F").NumberFormat = "m/d/yyyy;@"

        FormatString = "Managerial Employees Due a Salary Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output_BPS, FormatString, 6)
        xlWS_Output_BPS.Range("F:F").NumberFormat = "m/d/yyyy;@"

        FormatString = "Custodians Due a Salary Increase in " + ProcessMonth + Str(Year(ConfigurationManager.AppSettings("defaultDate")))
        FormatWorksheet(xlWS_Output_SCA, FormatString, 2)
        xlWS_Output_SCA.Range("F:F").NumberFormat = "m/d/yyyy;@"
        xlWS_Output_SCA.Range("K:K").NumberFormat = "m/d/yyyy;@"

        'Close Worksheets
        xlWB_Merged.Save()
        xlWB_Merged.Close()
        xlWS_Merged = Nothing
        xlWB_Merged = Nothing
        xlWB_Output.SaveAs(fileOutput.FileName)
        xlWB_Output.Close()
        xlWS_Output = Nothing
        xlWS_Output_AFU = Nothing
        xlWS_Output_PAA = Nothing
        xlWS_Output_SU7 = Nothing
        xlWS_Output_SCA = Nothing
        xlWS_Output_BPS = Nothing
        xlWB_Output = Nothing

        'Close Excel Application
        xlApp.Quit()
        xlApp = Nothing


        Windows.Forms.Cursor.Current = Cursors.Arrow

    End Sub

#End Region

#Region "Functions/Subs"

    Private Function GetYear(ByVal sDate As Object) As Integer

        If sDate Is Nothing Then
            Return 0
        Else
            Return CDate(sDate.value).Year

        End If

    End Function

    Private Function GetMonth(ByVal sDate As Object) As Integer

        If sDate Is Nothing Then
            Return 0
        Else
            Return CDate(sDate.Value).Month

        End If

    End Function

    Private Sub FormatWorksheet(Worksheet As Excel.Worksheet, Name As String, Sortby As Integer)

        Dim wsRange As Excel.Range

        Worksheet.Activate()
        wsRange = Worksheet.Range("A1", "K1")
        wsRange.Select()
        wsRange.Font.Bold = True

        ' Worksheet.Cells.Select()
        wsRange = Worksheet.Range("A1", "K65536")
        wsRange.Select()
        wsRange.Sort(Key1:=wsRange.Columns(Sortby), Order1:=Excel.XlSortOrder.xlAscending, Orientation:=Excel.XlSortOrientation.xlSortColumns, Header:=Excel.XlYesNoGuess.xlYes, OrderCustom:=1, MatchCase:=False)
        wsRange.Columns.AutoFit()

        With Worksheet.PageSetup
            .LeftHeader = ""
            .CenterHeader = _
            "Boston Public Schools" & Chr(10) & "Office of Human Resources" & Chr(10) & Name & Chr(10) & "By Alpha"
            .RightHeader = "&F"
            .LeftFooter = ""
            .CenterFooter = "Page &P of &N"
            .RightFooter = "&D &T"
            .LeftMargin = 36 ' Application.InchesToPoints(0.5)
            .RightMargin = 36 'Application.InchesToPoints(0.5)
            .TopMargin = 100 'Application.InchesToPoints(1.25)
            .BottomMargin = 72 'Application.InchesToPoints(1)
            .HeaderMargin = 36 'Application.InchesToPoints(0.5)
            .FooterMargin = 36 'Application.InchesToPoints(0.5)
            .PrintHeadings = False
            .PrintGridlines = True
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlLandscape
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperLetter
            '       .FirstPageNumber = xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 80
            .PrintTitleRows = "$1:$1"
            .PrintTitleColumns = ""
        End With

        If Worksheet.Name = "SCA" Then
            Worksheet.PageSetup.Zoom = 80
        End If


    End Sub

#End Region

End Class

