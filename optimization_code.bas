Attribute VB_Name = "Module1"
Sub OptimizeProcess()
    Dim wsInput As Worksheet
    Dim wsOptimization As Worksheet
    Dim wsReport As Worksheet
    Dim Temperature As Double
    Dim Pressure As Double
    Dim Catalyst As Double
    Dim ReactionTime As Double
    Dim Yield As Double
    Dim i As Integer
    Dim maxYield As Double
    Dim maxTemp As Double
    Dim maxPressure As Double
    Dim maxCatalyst As Double
    Dim maxTime As Double

    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsOptimization = ThisWorkbook.Sheets("Optimization")
    Set wsReport = ThisWorkbook.Sheets("Report")

    maxYield = 0
    
    ' Loop through SCADA data and calculate yield
    For i = 2 To 6
        Temperature = wsInput.Cells(i, 2).Value
        Pressure = wsInput.Cells(i, 3).Value
        Catalyst = wsInput.Cells(i, 4).Value
        ReactionTime = wsInput.Cells(i, 5).Value

        ' Example yield calculation
        Yield = (Temperature * Pressure) / (Catalyst * ReactionTime)
        
        ' Check if this is the maximum yield
        If Yield > maxYield Then
            maxYield = Yield
            maxTemp = Temperature
            maxPressure = Pressure
            maxCatalyst = Catalyst
            maxTime = ReactionTime
        End If
    Next i

    ' Output the optimal parameters and yield
    wsOptimization.Range("A1").Value = "Optimized Process Parameters"
    wsOptimization.Range("A2").Value = "Temperature"
    wsOptimization.Range("B2").Value = maxTemp
    wsOptimization.Range("A3").Value = "Pressure"
    wsOptimization.Range("B3").Value = maxPressure
    wsOptimization.Range("A4").Value = "Catalyst"
    wsOptimization.Range("B4").Value = maxCatalyst
    wsOptimization.Range("A5").Value = "Reaction Time"
    wsOptimization.Range("B5").Value = maxTime
    wsOptimization.Range("A6").Value = "Yield"
    wsOptimization.Range("B6").Value = maxYield

    ' Apply formatting to the Input and Optimization sheets
    Call FormatInputSheet(wsInput)
    Call FormatOptimizationSheet(wsOptimization)

    ' Generate Report
    Call GenerateReport(wsInput, maxTemp, maxPressure, maxCatalyst, maxTime, maxYield)
End Sub

Sub FormatInputSheet(ws As Worksheet)
    With ws
        ' Set column widths
        .Columns("A:E").AutoFit
        
        ' Apply borders
        .Range("A1:E6").Borders.LineStyle = xlContinuous

        ' Apply font and alignment
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E6").HorizontalAlignment = xlCenter
        .Range("A1:E6").VerticalAlignment = xlCenter

        ' Apply cell colors
        .Range("A1:E1").Interior.Color = RGB(0, 112, 192) ' Blue header
        .Range("A1:E1").Font.Color = RGB(255, 255, 255) ' White font for header
        .Range("A2:E6").Interior.Color = RGB(242, 242, 242) ' Light grey for data rows
    End With
End Sub

Sub FormatOptimizationSheet(ws As Worksheet)
    With ws
        ' Set column widths
        .Columns("A:B").AutoFit
        
        ' Apply borders
        .Range("A1:B6").Borders.LineStyle = xlContinuous

        ' Apply font and alignment
        .Range("A1:B1").Font.Bold = True
        .Range("A1:B6").HorizontalAlignment = xlCenter
        .Range("A1:B6").VerticalAlignment = xlCenter

        ' Apply cell colors
        .Range("A1").Interior.Color = RGB(0, 112, 192) ' Blue header
        .Range("A1").Font.Color = RGB(255, 255, 255) ' White font for header
        .Range("A2:B6").Interior.Color = RGB(242, 242, 242) ' Light grey for data rows
    End With
End Sub

Sub GenerateReport(wsInput As Worksheet, Temperature As Double, Pressure As Double, Catalyst As Double, ReactionTime As Double, Yield As Double)
    Dim wsReport As Worksheet
    Dim chartObj As ChartObject
    Dim chartRange As Range

    Set wsReport = ThisWorkbook.Sheets("Report")

    wsReport.Cells.Clear ' Clear previous report content

    ' Remove existing charts
    For Each chartObj In wsReport.ChartObjects
        chartObj.Delete
    Next chartObj

    ' Add headers and data to the Report sheet
    wsReport.Range("A1").Value = "Process Optimization Report"
    wsReport.Range("A2").Value = "Temperature"
    wsReport.Range("B2").Value = Temperature
    wsReport.Range("A3").Value = "Pressure"
    wsReport.Range("B3").Value = Pressure
    wsReport.Range("A4").Value = "Catalyst"
    wsReport.Range("B4").Value = Catalyst
    wsReport.Range("A5").Value = "Reaction Time"
    wsReport.Range("B5").Value = ReactionTime
    wsReport.Range("A6").Value = "Yield"
    wsReport.Range("B6").Value = Yield

    ' Apply formatting to the Report sheet
    Call FormatReportSheet(wsReport)

    ' Copy SCADA data to Report sheet for charting
    wsInput.Range("A1:E6").Copy Destination:=wsReport.Range("A8")

    ' Create the chart
    Set chartRange = wsReport.Range("A8:E13") ' Adjust the range as needed
    Set chartObj = wsReport.ChartObjects.Add(Left:=10, Top:=wsReport.Range("A15").Top, Width:=500, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=chartRange
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "SCADA Data"
        .Axes(xlCategory, xlPrimary).HasTitle = True
        .Axes(xlCategory, xlPrimary).AxisTitle.Text = "Reading #"
        .Axes(xlValue, xlPrimary).HasTitle = True
        .Axes(xlValue, xlPrimary).AxisTitle.Text = "Values"
        .Legend.Position = xlLegendPositionBottom
    End With

    ' Save Report as PDF
    wsReport.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.Path & "\Process_Optimization_Report.pdf"

    ' Switch to the Report sheet
    wsReport.Activate
End Sub

Sub FormatReportSheet(ws As Worksheet)
    With ws
        ' Set column widths
        .Columns("A:B").AutoFit
        
        ' Apply borders
        .Range("A1:B6").Borders.LineStyle = xlContinuous

        ' Apply font and alignment
        .Range("A1").Font.Bold = True
        .Range("A1:B6").HorizontalAlignment = xlCenter
        .Range("A1:B6").VerticalAlignment = xlCenter

        ' Apply cell colors
        .Range("A1").Interior.Color = RGB(0, 112, 192) ' Blue header
        .Range("A1").Font.Color = RGB(255, 255, 255) ' White font for header
        .Range("A2:B6").Interior.Color = RGB(242, 242, 242) ' Light grey for data rows
    End With
End Sub

