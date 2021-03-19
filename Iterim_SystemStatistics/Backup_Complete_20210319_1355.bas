Attribute VB_Name = "Module1"
Sub AssembleAndFormat()

    ControlWorkbook = "MacroAndStageAlternative_V1-2.xlsm"
    PickA = Sheets("Setup").Range("C5").Value
    PickB = Sheets("Setup").Range("C7").Value
    ChangePick = "Change"
    PickA_Abbreviation = Sheets("Setup").Range("E5").Value
    PickB_Abbreviation = Sheets("Setup").Range("E7").Value
    
    With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = False
        'Show the dialog box
        .Show
        'Store in fullpath variable
        ResultSetPath_str = .SelectedItems.Item(1)
    End With

    Workbooks.Open Filename:=ResultSetPath_str
    ResultSetWbName = ActiveWorkbook.Name
'''
   '*** RETRIEVE AND CONFORM DATASET DISPLAY TO REPORT CONFIGURATIONS ****
    'Retrieve AuditRouteList: requires row 1 shift and grouped columns
    Sheets("AuditRouteList").Select
    Sheets("AuditRouteList").Copy After:=Workbooks(ControlWorkbook).Sheets(2)
    Rows("1:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report and conditional formatting
    Range("A1").Value = "Route Audit Report"
    Range("A2").Value = "Cells highlighted golden have either no peak service or no Weekday schedule(s)."

    'Merge the grouped columns
    Range("B3:F3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
        .WrapText = True
    End With
    ActiveCell.Value = "Route in HASTUS Green Method"

    Range("G3:J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
        .WrapText = True
    End With
    ActiveCell.Value = "Route in Hours-Miles-Trips & Veh Stats"

    'Hard code replacement values for attributes not used elsewhere
    Range("B4").Value = "AM peak"
    Range("C4").Value = "Midday"
    Range("D4").Value = "PM peak"
    Range("E4").Value = "Night"
    Range("F4").Value = "Owl"
    Range("G4").Value = "weekday"
    Range("H4").Value = "Saturday"
    Range("I4").Value = "Sunday"
    Range("J4").Value = "Weekly"

    'Apply conditional formatting to routes with either no am or pm peak, or with no weekday service
    Range("A5").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=OR(B4=0, D4=0, G4=0)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False

    Rows("1:2").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Columns("B:J").AutoFit

    Windows(ResultSetWbName).Activate
''
    'Retrieve SvcStatsGar: requires row 1 shift and grouped columns
    Sheets("SvcStatsGar").Select
    Sheets("SvcStatsGar").Copy After:=Workbooks(ControlWorkbook).Sheets(3)
    Rows("1:2").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Service Statistics by Day of Week, Provider, and Garage"

    'Merge the grouped columns
    Range("D2:F2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Range("G2:I2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Range("J2:L2").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    Rows("1").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Windows(ResultSetWbName).Activate
''
    'Retrieve RteTripGar: requires row 1 shift and grouped columns
    Sheets("RteTripGar").Select
    Sheets("RteTripGar").Copy After:=Workbooks(ControlWorkbook).Sheets(4)
    Rows("1:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Report -  Hours and Trips by Route and Garage"
    Range("A2").Value = "Only contains routes with data from " & PickA & " " & PickB & " extracts."

    'Merge the grouped columns
    Columns("J:K").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("D3:K3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Columns("R:S").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("L3:S3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Columns("Z:AA").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("U3:AA3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    Rows("1:2").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Windows(ResultSetWbName).Activate
'''
    'Retrieve RteTripPvdr: requires row 1 shift and grouped columns
    Sheets("RteTripPvdr").Select
    Sheets("RteTripPvdr").Copy After:=Workbooks(ControlWorkbook).Sheets(5)
    Rows("1:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Report -  Hours and Trips by Route and Provider"
    Range("A2").Value = "Only contains routes with data from " & PickA & " " & PickB & " extracts."

    'Merge the grouped columns
    Columns("J:K").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("D3:K3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Columns("R:S").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("L3:S3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Columns("Z:AA").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("U3:AA3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    Rows("1:2").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Windows(ResultSetWbName).Activate
'''
    'Retrieve RteTrips: requires row 1 shift and grouped columns
    Sheets("RteTrips").Select
    Sheets("RteTrips").Copy After:=Workbooks(ControlWorkbook).Sheets(6)
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Report -  Hours and Trips by Route"
    Range("A2").Value = "Only contains routes with data from " & PickA & " " & PickB & " extracts."

    'Merge the grouped columns
    Columns("I:J").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("C3:J3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Columns("Q:R").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("K3:R3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Columns("Y:Z").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With
    Range("S3:Z3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    Rows("1:2").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Windows(ResultSetWbName).Activate
''
    'Retrieve PlatHrsGar: requires row 1 shift and grouped columns
    Sheets("PlatHrsGar").Select
    Sheets("PlatHrsGar").Copy After:=Workbooks(ControlWorkbook).Sheets(7)
    Rows("1:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Report -  Platform hours by Garage"
    Range("A2").Value = "Use the filter on Garage to see/hide the Subtotal rows for each Provider."

    'Merge the grouped columns
    Range("C3:E3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Range("F3:H3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Range("I3:K3").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    Rows("1:2").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Range("B4").Select
    Selection.End(xlDown).Select
    DeleteRow = ActiveCell.Row
    Rows("" & DeleteRow & "").Select
    Selection.Delete Shift:=xlUp

    Windows(ResultSetWbName).Activate
'''
'    Retrieve PeakBusType: requires row 1 shift and grouped columns
    Sheets("PeakBusType").Select
    Sheets("PeakBusType").Copy After:=Workbooks(ControlWorkbook).Sheets(8)
    Rows("1:4").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Peak Bus Requirements using HASTUS Green Method"
    Range("A2").Value = "Use the filters see/hide the Subtotal rows for each Vehicle Type, Block Garage, and Provider."

    'The grouping pre-merge needs to follow a different sequence than previous tabs for creating Period groups
    Columns("H:J").Group
    Columns("N:S").Group

    'Now merge with labels for time period; again these are hard code based on known output.
    Range("E4:G4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "AM peak period"

    Range("H4:J4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "Mid-day period"

    Range("K4:M4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "PM peak period"

    Range("N4:P4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "Night period"

    Range("Q4:S4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "Owl period"

       Range("T4:V4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "Max AM/PM Peak Periods"

    Range("W4:Y4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = "Max All Periods"

    Range("E5").Value = PickA_Abbreviation
    Range("F5").Value = PickB_Abbreviation
    Range("G5").Value = ChangePick
    Range("E5:G5").Select
    Selection.Copy
    Range(Selection, Selection.End(xlToRight)).Select
    ActiveSheet.Paste

    Columns("H:J").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("N:S").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Rows("1:3").Select
    Selection.Rows.Group
    With Selection
        .Rows.Group
        .EntireRow.Hidden = False
    End With

    Windows(ResultSetWbName).Activate
'''
    'Retrieve WindowLocalMinMax: requires row 1 shift and grouped columns
    MsgBox ("Built from " & PickA_Abbreviation & " and " & PickB_Abbreviation & ". Check grouping with raw data!")
    Sheets("WindowLocalMinMax").Select
    Sheets("WindowLocalMinMax").Copy After:=Workbooks(ControlWorkbook).Sheets(9)
    Rows("1:5").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

    'Label report
    Range("A1").Value = "Peak Bus Requirements using local Minimums and Maximums within 3-hour intervals"
    Range("A2").Value = "Use the filters see/hide the Subtotal rows for each Vehicle Type, Block Garage, and Provider."
    Range("A3").Value = "Extensive document outlining applied to easily view " & PickA_Abbreviation & ", " & PickB_Abbreviation & ", and " & ChangePick & ". Open to revision!"

    'The grouping pre-merge needs to follow a different sequence than previous tabs for creating Period groups
    'Outside intervals similar to the AM/PM Peak periods
    Columns("E:H").Group
    Columns("K:N").Group
    Columns("Q:AB").Group
    Columns("AE:AH").Group
    Columns("AK:AV").Group
    Columns("AY:BC").Group
    Columns("BF:BL").Group
    'Pick B so that Pick A and Change show
    Columns("Y:AS").Group 'Pick B

    Columns("E:H").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("K:N").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("Q:AB").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("AE:AH").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("AK:AV").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("AY:BB").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Columns("BE:BL").Select
    With Selection
        .Columns.Group
        .EntireColumn.Hidden = True
    End With

    Range("E4:X4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickA

    Range("Y4:AR4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = PickB

    Range("AS4:BL4").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .MergeCells = True
    End With
    ActiveCell.Value = ChangePick

    'Have not found an efficiency for dynamically retrieving the names of the interval with the merge!
    Range("E5:F5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("G5:H5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("I5:J5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("K5:L5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("M5:N5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("O5:P5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("Q5:R5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("S5:T5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("U5:V5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("W5:X5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("Y5:Z5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AA5:AB5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AC5:AD5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AE5:AF5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AG5:AH5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AI5:AJ5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AK5:AL5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AM5:AN5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AO5:AP5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AQ5:AR5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AS5:AT5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AU5:AV5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AW5:AX5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("AY5:AZ5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BA5:BB5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BC5:BD5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BE5:BF5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BG5:BH5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BI5:BJ5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("BK5:BL5").Select
    With Selection
        .MergeCells = True
        .HorizontalAlignment = xlCenter
    End With
    ActiveCell.FormulaR1C1 = "=""["" & LEFT(RIGHT(R[1]C,5),2) & "":00, "" & RIGHT(RIGHT(R[1]C,6), 2) & "":00)"""

    Range("E5:BL5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False

    Range("E6").Value = "Min"
    Range("F6").Value = "Max"
    Range("E6:F6").Select
    Selection.Copy
    Range("G6:BL6").Select
    ActiveSheet.Paste


    Windows(ResultSetWbName).Activate
'''
    'Retrieve VehTaskReq: requires row 1 shift and grouped columns
    Sheets("VehTaskReq").Select
    Sheets("VehTaskReq").Copy After:=Workbooks(ControlWorkbook).Sheets(10)
    Rows("1:3").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Application.CutCopyMode = False
'
'    'Label report
    Range("A1").Value = "Vehicle Tasks Summary"
    Range("A2").Value = "Use the filters see/hide the Subtotal rows for each Vehicle Type and Block Garage." & _
    "Uses count distinct method and sums of count distinct for subtotals."

    'Label Picks with abbreviations
    Range("E3").Value = PickA_Abbreviation
    Range("F3").Value = PickB_Abbreviation
    Range("G3").Value = ChangePick
'''
    Windows(ResultSetWbName).Close
    Windows(ControlWorkbook).Activate
'''
'   '*** BEGIN VARIABLE CLEANUP ****
' Assumes starting position in [AuditRouteList] tab where Workbook has 2 setup tabs
    
    Sheets("Setup").Activate
    StartRow = 20
    Sheets("Setup").Range("B" & StartRow).Select
    
    For i = 1 To 9
    SheetName = Range("B" & StartRow).Value
    
    Worksheets("" & SheetName & "").Activate
    
    ' *** Replace Variable Names ***
    ' Some use LookAt:=xlPart and others use LookAt:=xlWhole on whether to match with wildcard or entire cell value, respectively

    'Dimensions
    'sched_type ==> "Day of week"
    Cells.Replace What:="sched_type", Replacement:="Day of week", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'gargroup ==> "Provider"
    Cells.Replace What:="gargroup", Replacement:="Provider", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'garage ==> "Garage"
    Cells.Replace What:="garage", Replacement:="Garage", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'block_garage ==> "Blk Garage"
    Cells.Replace What:="block_garage", Replacement:="Blk Garage", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'route ==> "Route"
    Cells.Replace What:="route", Replacement:="Route", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'veh-task_veh_group ==> "Veh"
    Cells.Replace What:="veh_task_veh_group", Replacement:="Veh", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'Route Audit Measures
    'None -- placeholder

    'Hours-Miles-Trips Measures
    'insrvhrs ==> In-srv Hrs
    Cells.Replace What:="*insrvhrs*", Replacement:="In-Srv Hrs", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'plathrs_ ==> Plat hrs
    Cells.Replace What:="plathrs_*", Replacement:="Plat hrs", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*plathrs* ==> Plat hrs
    Cells.Replace What:="change_plathrs*", Replacement:="Plat hrs", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*trips_ ==> Trips
    Cells.Replace What:="*trips_", Replacement:="Trips", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*change_trips ==> Trips
    Cells.Replace What:="change_trips", Replacement:="Trips", LookAt:= _
        xlWhole, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    'Block Count Measures
    '*am_count* ==> AM peak
    Cells.Replace What:="*am_count*", Replacement:="AM peak", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*mid_count* ==> Mid day
    Cells.Replace What:="*mid_count*", Replacement:="Mid day", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*pm_count* ==> PM peak
    Cells.Replace What:="*pm_count*", Replacement:="PM peak", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*night_count* ==> Night
    Cells.Replace What:="*night_count*", Replacement:="Night", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*owl_count* ==> Night
    Cells.Replace What:="*owl_count*", Replacement:="Owl", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    '*ctd_veh_task* ==> Veh task count
    Cells.Replace What:="*ctd_veh_task*", Replacement:="Veh task count", LookAt:= _
        xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2

    ' *** Autofit to New Column Width ***
    Application.Goto Reference:="R6C2"
    Selection.End(xlToRight).Select

    Dim LastColumnNumber As Integer
        LastColumnNumber = ActiveCell.Column
    Dim LastColumnLetter As String
        LastColumnLetter = Split(Cells(1, LastColumnNumber).Address, "$")(1)

    Columns("B:" & LastColumnLetter & "").AutoFit
    
    Sheets("Setup").Activate
    ActiveCell.Offset(1, 0).Select
    If ActiveCell.Value <> "" Then
        StartRow = ActiveCell.Row
        Else
        dummy = 1
    End If
   
Next i

' *** ADD COSMETIC TOUCHES TO OVERCOME DEFAULTS FROM VBA ****
' These are the finishing touches that are custom to each Sheet

    Sheets("AuditRouteList").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D20").Value = "AuditRouteList"
    
    Sheets("SvcStatsGar").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Columns("D:L").Select
    Selection.ColumnWidth = 8.2
    Rows("2:2").Font.Underline = xlUnderlineStyleSingle
    Rows("3:3").AutoFilter
    Sheets("Setup").Range("D21").Value = "SvcStatsGar"

    Sheets("RteTripGar").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D22").Value = "RteTripGar"

    Sheets("RteTripPvdr").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D23").Value = "RteTripPvdr"

    Sheets("RteTrips").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D24").Value = "RteTrips"
    
    Sheets("PlatHrsGar").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Columns("C:K").Select
    Selection.ColumnWidth = 6.3
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D25").Value = "PlatHrsGar"
    
    Sheets("PeakBusType").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Rows("4:4").Font.Underline = xlUnderlineStyleSingle
    Rows("5:5").AutoFilter
    Sheets("Setup").Range("D26").Value = "PeakBusType"

    
    Sheets("WindowLocalMinMax").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:3").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Columns("E:BL").Select
    Selection.ColumnWidth = 5.86
    Rows("4:5").Font.Underline = xlUnderlineStyleSingle
    Rows("6:6").AutoFilter
    Sheets("Setup").Range("D27").Value = "WindowLocalMinMax"
    
    Sheets("VehTaskReq").Select
    Rows("1:1").Select
    Selection.Font.Bold = True
    Rows("2:2").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
    End With
    Rows("3:3").Font.Underline = xlUnderlineStyleSingle
    Rows("4:4").AutoFilter
    Sheets("Setup").Range("D28").Value = "VehTaskReq"
    
End Sub


































