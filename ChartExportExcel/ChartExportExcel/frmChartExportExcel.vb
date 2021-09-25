#Region "ABOUT"
' / --------------------------------------------------------------------------
' / Developer : Mr.Surapon Yodsanga (Thongkorn Tubtimkrob)
' / eMail : thongkorn@hotmail.com
' / URL: http://www.g2gnet.com (Khon Kaen - Thailand)
' / Facebook: https://www.facebook.com/g2gnet (For Thailand)
' / Facebook: https://www.facebook.com/commonindy (Worldwide)
' / More Info: http://www.g2gnet.com/webboard
' /
' / Purpose: Create data into DataGridView for make the Chart and Export data to Excel @Run Time.
' / Microsoft Visual Basic .NET (2010) + MS Excel 14.0
' /
' / This is open source code under @CopyLeft by Thongkorn Tubtimkrob.
' / You can modify and/or distribute without to inform the developer.
' / --------------------------------------------------------------------------
#End Region

Imports System.Windows.Forms.DataVisualization.Charting
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO

Public Class frmChartExportExcel

    Private Sub frmPlotGraph_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Label1.Text = ""
        '// Chart Type
        '// https://docs.microsoft.com/en-us/dotnet/api/system.windows.forms.datavisualization.charting.seriescharttype?view=netframework-4.7.2
        cmbChartType.Items.Add("Column")
        cmbChartType.Items.Add("Line")
        cmbChartType.Items.Add("Point")
        cmbChartType.SelectedIndex = 0

        '// Setup DataGridView
        Call InitializeGrid()
        '// Generate Data
        Call btnGenData_Click(sender, e)
        '//
    End Sub

    Private Sub btnGenData_Click(sender As System.Object, e As System.EventArgs) Handles btnGenData.Click
        Call FillData()
        Call btnChart_Click(sender, e)
    End Sub

    ' / --------------------------------------------------------------------------
    Private Sub FillData()
        Dim dt As New DataTable
        dt.Columns.Add("Item")
        dt.Columns.Add("Value1")
        dt.Columns.Add("Value2")
        dt.Columns.Add("Value3")
        dt.Columns.Add("Average")
        Dim RandomClass As New Random()
        Dim MyMonth As String() = New String() {("ม.ค."), ("ก.พ."), ("มี.ค."), ("เม.ย."), ("พ.ค."), ("มิ.ย."), ("ก.ค."), ("ส.ค."), ("ก.ย."), ("ต.ค."), ("พ.ย."), ("ธ.ค.")}
        For i As Byte = 0 To 11
            Dim dr As DataRow = dt.NewRow()
            dr(0) = MyMonth(i)
            dr(1) = FormatNumber(RandomClass.Next(100, 1000) + RandomClass.NextDouble(), 2)
            dr(2) = FormatNumber(RandomClass.Next(100, 1000) + RandomClass.NextDouble(), 2)
            dr(3) = FormatNumber(RandomClass.Next(100, 1000) + RandomClass.NextDouble(), 2)
            '// CDbl = Convert to Double
            dr(4) = Format((CDbl(dr(1)) + CDbl(dr(2)) + CDbl(dr(3))) / 3, "#,##0.00")
            dt.Rows.Add(dr)
        Next
        dgvData.DataSource = dt
        Label1.Text = "Total : " & dt.Rows.Count.ToString("#,##") & " Records."
    End Sub

    ' / --------------------------------------------------------------------------------
    '// การตั้งค่าเริ่มต้นให้กับตารางกริดในแบบ @Run Time
    Private Sub InitializeGrid()
        With dgvData
            .RowHeadersVisible = False
            .AllowUserToAddRows = False
            .AllowUserToDeleteRows = False
            .AllowUserToResizeRows = False
            .MultiSelect = False
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect
            .ReadOnly = True
            .Font = New Font("Tahoma", 9)
            '/ จัดความกว้างของแต่ละหลัก
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            .AutoResizeColumns()
            '/ Adjust Header Styles
            With .ColumnHeadersDefaultCellStyle
                .BackColor = Color.Navy
                .ForeColor = Color.White
                .Font = New Font("Tahoma", 9, FontStyle.Bold)
            End With
        End With
    End Sub

    ' / --------------------------------------------------------------------------------
    Private Sub btnChart_Click(sender As System.Object, e As System.EventArgs) Handles btnChart.Click
        If dgvData.RowCount = 0 Then Return
        '//
        Chart1.Titles.Clear()
        '// Clear Legends & Series
        Chart1.Series.Clear()
        Chart1.Legends.Clear()
        '// Clear Points
        For Each xSeries In Chart1.Series
            xSeries.Points.Clear()
        Next
        '// Create Legends
        Dim Legend1 As Legend = New Legend()
        Legend1.Name = "Legend1"
        Chart1.Legends.Add(Legend1)
        '// Create Series
        Dim Series1 As Series = New Series()
        Dim Series2 As Series = New Series()
        Dim Series3 As Series = New Series()
        Dim Series4 As Series = New Series()
        '// Add Series
        With Chart1
            .Series.Add(Series1)
            .Series.Add(Series2)
            .Series.Add(Series3)
            .Series.Add(Series4)
        End With
        '//
        Try
            Dim myFont As New Font("Tahoma", 14, FontStyle.Bold)
            Chart1.Titles.Add("Sample Chart")
            Chart1.Titles(0).Font = myFont
            '// Define Chart Type.
            For Each xSeries In Chart1.Series
                Select Case cmbChartType.Text
                    Case "Column"
                        xSeries.ChartType = SeriesChartType.Column
                    Case "Line"
                        xSeries.ChartType = SeriesChartType.Line
                    Case "Point"
                        xSeries.ChartType = SeriesChartType.Point
                End Select
                '// Show Legend
                If chkShowLegend.Checked Then
                    xSeries.IsVisibleInLegend = True
                Else
                    xSeries.IsVisibleInLegend = False
                End If
                '// View the value of a chart point on mouse over.
                xSeries.ToolTip = "#VAL{0.00}"
            Next
            '//
            With Chart1
                .Series(0).LegendText = "Value 1"
                .Series(1).LegendText = "Value 2"
                .Series(2).LegendText = "Value 3"
                .Series(3).LegendText = "Average"
            End With
            '// Show all Axis Label
            Chart1.ChartAreas(0).AxisX.Interval = 1
            '//
            With Chart1.ChartAreas("ChartArea1")
                .AxisX.MajorGrid.LineWidth = 1
                .AxisY.MajorGrid.LineWidth = 1
            End With

            '// LOOP DISPLAY
            For CountData As Integer = 0 To dgvData.Rows.Count - 1
                With Chart1
                    '// 0 = Month, 1 = Value1, 2 = Value2, 3 = Value3, 4 = Average
                    .Series(0).Points.AddXY(dgvData.Item(0, CountData).Value, dgvData.Item(1, CountData).Value)
                    .Series(1).Points.AddXY(dgvData.Item(0, CountData).Value, dgvData.Item(2, CountData).Value)
                    .Series(2).Points.AddXY(dgvData.Item(0, CountData).Value, dgvData.Item(3, CountData).Value)
                    .Series(3).Points.AddXY(dgvData.Item(0, CountData).Value, dgvData.Item(4, CountData).Value)
                End With
            Next
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    ' / --------------------------------------------------------------------------------
    '// Export data to Excel.
    Private Sub btnExportExcel_Click(sender As System.Object, e As System.EventArgs) Handles btnExportExcel.Click
        '// ไม่มีข้อมูลในตารางกริด ก็สั่งให้เด้งหนีออกไป
        If dgvData.Rows.Count = 0 Then Exit Sub
        '// พิจารณาการเลือกใช้ชนิดข้อมูล (Data Type) ให้เหมาะสม
        Dim MaxRow As Integer, MaxCol As Short
        Dim nRow As Integer, nCol As Short
        Dim xlsApp As New Excel.Application
        Dim xlsWorkBook As Excel.Workbook = xlsApp.Workbooks.Add
        Dim xlsWorkSheet As Excel.Worksheet = CType(xlsWorkBook.Worksheets(1), Excel.Worksheet)
        '// S T A R T
        Try
            xlsApp.Visible = True
            '// หาค่าจำนวนแถว
            MaxRow = dgvData.RowCount
            '// หาค่าจำนวนหลัก
            MaxCol = dgvData.Columns.Count - 1
            With xlsWorkSheet
                .Cells.Select()
                .Cells.Delete()
                '// Header
                For nCol = 0 To MaxCol
                    .Cells(1, nCol + 1).Value = dgvData.Columns(nCol).HeaderText
                Next nCol
                '// ไล่ตามจำนวนแถว
                For nRow = 0 To MaxRow - 1
                    For nCol = 0 To MaxCol
                        .Cells(nRow + 2, nCol + 1).value = dgvData.Rows(nRow).Cells(nCol).Value
                    Next nCol   '// Nested Loop
                    '// หากชุดคำสั่งที่อยู่ใน Nested Loop 
                    '// การให้ตัวแปรต่อท้าย Next จะช่วยให้เรารู้ว่ามันอยู่ใน Loop ไหน
                Next nRow

                '// กำหนดรูปแบบใน WorkSheet
                .Rows("1:1").Font.FontStyle = "Bold"
                .Rows("1:1").Font.Size = 10
                '//
                .Cells.Columns.AutoFit()
                .Cells.Select()
                .Cells.EntireColumn.AutoFit()
                .Cells(1, 1).Select()
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        xlsApp = Nothing
        '// Release memory.
        ReleaseObject(xlsApp)
        ReleaseObject(xlsWorkBook)
        ReleaseObject(xlsWorkSheet)
    End Sub

    ' / ------------------------------------------------------------------
    '// คืนหน่วยความจำกลับคืนสู่ระบบ
    Public Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            '// GC = Garbage Collection
            GC.Collect()
        End Try
    End Sub

    Private Sub frmPlotGraph_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Me.Dispose()
        GC.SuppressFinalize(Me)
        Application.Exit()
    End Sub

    Private Sub btnClose_Click(sender As System.Object, e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    '// เวลากด Enter ในแถวของตารางกริด จะไม่เกิดการเลื่อนตำแหน่ง
    Private Sub dgvData_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles dgvData.KeyDown
        If e.KeyCode = Keys.Enter Then e.SuppressKeyPress = True
    End Sub

    Private Sub cmbChartType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbChartType.SelectedIndexChanged
        Call btnChart_Click(sender, e)
    End Sub

    Private Sub chkShowLegend_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkShowLegend.CheckedChanged
        Call btnChart_Click(sender, e)
    End Sub

    '// Copy to Clipboard
    Private Sub btnCopyClipboard_Click(sender As System.Object, e As System.EventArgs) Handles btnCopyClipboard.Click
        Using ms As MemoryStream = New MemoryStream()
            Chart1.SaveImage(ms, ChartImageFormat.Bmp)
            Dim bm As Bitmap = New Bitmap(ms)
            Clipboard.SetImage(bm)
        End Using
    End Sub
End Class

