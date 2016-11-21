VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmGraph 
   Caption         =   "Graphical and Tabular representations"
   ClientHeight    =   8100
   ClientLeft      =   -1185
   ClientTop       =   795
   ClientWidth     =   14040
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   14040
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "&Export to Excel"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Frame fraMain 
      Height          =   6615
      Left            =   360
      TabIndex        =   0
      Top             =   -120
      Width           =   13335
      Begin VB.Frame fraSource 
         Caption         =   "Please Choose Data Source"
         Height          =   1335
         Left            =   8640
         TabIndex        =   17
         Top             =   240
         Width           =   3735
         Begin VB.CheckBox chckSource 
            Caption         =   "Local"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   20
            Top             =   960
            Width           =   975
         End
         Begin VB.CheckBox chckSource 
            Caption         =   "Regional"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox chckSource 
            Caption         =   "National"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   18
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame fraSec 
         BorderStyle     =   0  'None
         Caption         =   "Please choose the categories you want to compare"
         Height          =   3855
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   4455
         Begin VB.ListBox lstSecondry 
            Height          =   2985
            Left            =   480
            MultiSelect     =   1  'Simple
            TabIndex        =   8
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label lblList 
            Caption         =   "Please choose the categories you want to compare"
            Height          =   255
            Left            =   0
            TabIndex        =   14
            Top             =   0
            Width           =   4455
         End
      End
      Begin VB.Frame fraGraph 
         BorderStyle     =   0  'None
         Height          =   4455
         Left            =   4920
         TabIndex        =   12
         Top             =   1800
         Width           =   4335
         Begin VB.OLE oleGraph 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            BorderStyle     =   0  'None
            Class           =   "Excel.Chart.8"
            Height          =   3495
            Left            =   360
            SizeMode        =   1  'Stretch
            TabIndex        =   13
            Top             =   600
            Width           =   3735
         End
      End
      Begin VB.Frame fraChart 
         Caption         =   "Choose Comparison Subject"
         Height          =   975
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4335
         Begin VB.ComboBox cmbField 
            Height          =   315
            ItemData        =   "frmGraph.frx":0000
            Left            =   2400
            List            =   "frmGraph.frx":0002
            TabIndex        =   16
            Text            =   "Choose comparison field"
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox cmbGraph 
            Height          =   315
            ItemData        =   "frmGraph.frx":0004
            Left            =   480
            List            =   "frmGraph.frx":0006
            TabIndex        =   11
            Text            =   "Choose comparison area"
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblChart 
            Height          =   495
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame fraCat 
         Caption         =   "Please Choose Categories"
         Height          =   1335
         Left            =   4680
         TabIndex        =   4
         Top             =   240
         Width           =   3735
         Begin VB.ComboBox cmbCat 
            Height          =   315
            ItemData        =   "frmGraph.frx":0008
            Left            =   240
            List            =   "frmGraph.frx":0015
            TabIndex        =   5
            Text            =   "Main Categories (White, BME, Not Stated)"
            Top             =   840
            Width           =   3375
         End
      End
      Begin VB.Frame fraMCat 
         BorderStyle     =   0  'None
         Caption         =   "Please choose the categories you want to compare"
         Height          =   3855
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   4455
         Begin VB.ListBox lstMain 
            Height          =   2985
            Left            =   480
            MultiSelect     =   1  'Simple
            TabIndex        =   10
            Top             =   480
            Width           =   3735
         End
         Begin VB.Label Label1 
            Caption         =   "Please choose the categories you want to compare"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   4455
         End
      End
   End
   Begin VB.Image imgLogo 
      Height          =   1020
      Left            =   0
      Picture         =   "frmGraph.frx":00EB
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1155
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rGraph As Excel.Range
Dim xlSeries As Excel.Series
Dim oWBGraphs As Excel.Workbook
Dim oWSGraphs As Excel.Worksheet

Dim intCat As Integer '0 main,1 primary, 2 secondary
Dim arrEthnicities(5) As Integer
Dim strSheet As String
Dim intSource(3) As Integer
Dim intFields As Integer
Dim intStart As Integer
Dim intField As Integer

'strSht is the source sheet selected in cmbGraph
'intSource() array of the check boxes indeces to indicate which source (National,Regional,Local)
'intField indicate the col in the sheet for example FOR "Arabic" IN LANGUAGES 0
Sub plotGraphM()
    Dim i As Integer
    Dim j As Integer
    'source sheet
     Set oWS = oWB.Worksheets.Item(strSheet)
     'oWS.Cells(intStart, 3 + intField).Value
     
     'destination sheet
    If CheckPath(App.Path & "\res.xls") Then Kill App.Path & "\res.xls"
    If CheckPath(App.Path & "\res.xlsx") Then Kill App.Path & "\res.xlsx"
    Set oWBGraphs = oExcel.Workbooks.Open(App.Path & "\Res")
    Set oWSGraphs = oWBGraphs.Worksheets.Item(1)
    Set rGraph = oWSGraphs.Range(oWSGraphs.Cells(1, 1), oWSGraphs.Cells(6, UBound(arrEthnicities) + 1))
    rGraph(1, 1).Value = "Year"
    rGraph(2, 1).Value = oWS.Range("b2").Value
    rGraph(3, 1).Value = oWS.Range("b30").Value
    rGraph(4, 1).Value = oWS.Range("b60").Value
    rGraph(5, 1).Value = oWS.Range("b90").Value
    rGraph(6, 1).Value = oWS.Range("b120").Value
    
    rGraph.Range("a1:a6").NumberFormat = "@"
    Set oChart = oWBGraphs.Charts.Add
    With oChart
    'chart title
        .HasTitle = True
        .ChartTitle.Characters.Text = oWS.Name
        
    'legend position
    .HasLegend = False
     .ChartType = xlLineMarkers
        
        .Axes(xlValue, xlPrimary).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).MajorGridlines.Border.LineStyle = xlDot
        .Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = 15
        .SetSourceData Source:=rGraph, PlotBy:=xlColumns
        .Axes(xlCategory).HasTitle = True
       .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "year"
        .Location Where:=xlLocationAsNewSheet, Name:="Chr"
        For i = 0 To UBound(arrEthnicities) - 1 'category white,.....
            rGraph(1, 2 + i).Value = oWS.Range("a" & arrEthnicities(i)).Value & " " & oWS.Range("b" & arrEthnicities(i)).Value
            For j = 0 To Year(Date) - 2004
                
                rGraph(j + 2, 2 + i).Value = oWS.Range("c" & (30 * j + (7 + i))).Value
                rGraph(j + 2, 2 + i).NumberFormat = "0.00"
            Next j
            .SeriesCollection(i + 1).XValues = rGraph.Range("a2:a6")
            .SeriesCollection(i + 1).Values = rGraph.Range(rGraph.Cells(2, 2 + i), rGraph.Cells(6, 2 + i))
            
            .SeriesCollection(i + 1).Points(.SeriesCollection(i + 1).Points.Count - 1).ApplyDataLabels AutoText:=True, _
        LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
        ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False
            .SeriesCollection(i + 1).Name = rGraph(1, 2 + i).Value
            
        With .SeriesCollection(i + 1).Border
           .ColorIndex = 5    'q2
           .Weight = xlThick
            .LineStyle = xlContinuous
            
        End With
            With .SeriesCollection(i + 1)
                .MarkerBackgroundColorIndex = xlAutomatic
                .MarkerForegroundColorIndex = xlAutomatic
                .MarkerStyle = xlAutomatic
                .Smooth = False
                .MarkerSize = 5
                .Shadow = False
            End With
            .ChartArea.Font.Size = 12
            .PlotArea.Interior.ColorIndex = 2

            .PlotArea.Interior.PatternColor = 3
            '.PlotArea.Height = 200000
          
            'ActiveChart.SeriesCollection(1).Points(11).
    ' ows.range("a1").Characters.
    Next i

        
   
    End With
   ' oChart.HeightPercent = 80
 
End Sub

Private Sub cmbCat_click()
    'main
    If cmbCat.ListIndex = 0 Then
        fraMCat.Visible = False
        fraMCat.Enabled = False
        fraSec.Visible = False
        fraSec.Enabled = False
        intWidth = fraGraph.Width
        intLeft = fraGraph.Left
        fraGraph.Width = fraMain.Width - 2 * fraChart.Left
        fraGraph.Left = fraChart.Left
        fraGraph.Top = fraChart.Top + fraChart.Height
     
    ElseIf cmbCat.ListIndex = 1 Then
        fraMCat.Visible = True
        fraMCat.Enabled = True
        fraSec.Visible = False
        fraSec.Enabled = False
        fraGraph.Width = intWidth
        fraGraph.Left = intLeft
        oleGraph.Left = 0
        oleGraph.Top = 0
        oleGraph.Width = fraGraph.Width
        oleGraph.Height = fraGraph.Height
    ElseIf cmbCat.ListIndex = 2 Then
        fraMCat.Visible = False
        fraMCat.Enabled = False
        fraSec.Visible = True
        fraSec.Enabled = True
        fraGraph.Width = intWidth
        fraGraph.Left = intLeft
 
    End If
        oleGraph.Left = 0
        oleGraph.Top = 0
        oleGraph.Width = 0.95 * fraGraph.Width
        oleGraph.Height = fraGraph.Height
oleGraph.Enabled = False
On Error Resume Next
End Sub

Private Sub cmbField_click()
    intField = cmbField.ListIndex
End Sub

Private Sub cmbGraph_click()
 strSheet = oWB.Worksheets.Item(cmbGraph.List(cmbGraph.ListIndex)).Name
 cmbField.Clear

Dim intX As Integer
    Set oWS = oWB.Worksheets.Item(cmbGraph.Text)
    intFields = oWS.Cells(1, 1).Value
    intStart = oWS.Cells(1, 2).Value
    If intFields = 1 Then intFields = 2
    For intX = 3 To intFields + 1
        cmbField.AddItem oWS.Cells(intStart - 1, intX).Value
    Next intX
    cmbField.Text = cmbField.List(0)
End Sub

Private Sub cmdExcel_Click()
    ' Sets the Dialog Title to Save File
CommonDialog1.DialogTitle = "Save Excel File"

' Sets the File List box to Text File and All Files
CommonDialog1.Filter = "Excel Files 97-2003 (*.xls)|*.xls|Excel Files 2007(*.xlsx)|*.xlsx|AllFiles (*.*)|*.*"

' Set the default files type to Text File
CommonDialog1.InitDir = Environ$("USERPROFILE") & "\Desktop\"

' Sets the flags - Hide Read only, prompt to overwrite, and path must exist
CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNOverwritePrompt _
+ cdlOFNPathMustExist

' Set dialog box so an error occurs if the dialogbox is cancelled
CommonDialog1.CancelError = True

' Enables error handling to catch cancel error
On Error Resume Next
' display the dialog box
CommonDialog1.ShowSave
If Err Then
    ' This code runs if the dialog was cancelled
    MsgBox "Dialog Cancelled"
    Exit Sub
End If
FileCopy App.Path & "\agregateData.xls", CommonDialog1.FileName
Dim e As Excel.Application
Set e = New Excel.Application
 e.Workbooks.Open (CommonDialog1.FileName)
 e.Visible = True
End Sub

Private Sub cmdExit_Click()
    frmMain.Show
    Unload Me
End Sub

Private Sub Form_Load()
Dim i As Integer
intStart = 6
intFields = 1
intField = 1
    chckSource(0).Value = 1
    Set oExcel = New Excel.Application
    oExcel.Visible = True
    Set oWB = oExcel.Workbooks.Open(App.Path & "\agregateData.xls", False)
    formatControls
    loadLists

    intWidth = fraGraph.Width
    intLeft = fraGraph.Left
    cmbCat.ListIndex = 2
    cmbCat.Text = cmbCat.List(2)
    'ReDim arrEthnicities(0 To 3)
    arrEthnicities(0) = 7
    arrEthnicities(1) = 9
    arrEthnicities(2) = 10
    arrEthnicities(3) = 3
    arrEthnicities(4) = 8
    For i = 0 To 4
        lstSecondry.Selected(arrEthnicities(i)) = True
    Next i
    getComparison
    plotChart
    strSheet = oWB.Worksheets.Item(1).Name
    oleGraph.Class = "Excel.Chart8.0"

    oleGraph.CreateLink oWB.Path & "\" & oWB.Name, "chr"

    oleGraph.Update
    cmbField.Text = "Persons"

    'cmbCat.ListIndex = 0
    
Cleanup:
   ' Set oWS = Nothing
    'If Not oWB Is Nothing Then oWB.Close False
    'Set oWB = Nothing
    'oExcel.Quit
    'Set oExcel = Nothing
End Sub


'REGIONAL STARTS AT 2+LENGTH
'LOCAL STARTS AS 2+2*LENGTH
Private Sub plotChart()
'create a new workbook to store all graphs generated and data arranged
    
    Dim i As Integer
    Dim j As Integer
  
    oWB.Worksheets.Add
    Set rGraph = oWB.ActiveSheet.Range(oWB.ActiveSheet.Cells(1, 1), oWB.ActiveSheet.Cells(6, UBound(arrEthnicities) + 1))
    rGraph(1, 1).Value = "Year"
    rGraph(2, 1).Value = oWS.Range("b2").Value
    rGraph(3, 1).Value = oWS.Range("b30").Value
    rGraph(4, 1).Value = oWS.Range("b60").Value
    rGraph(5, 1).Value = oWS.Range("b90").Value
    rGraph(6, 1).Value = oWS.Range("b120").Value
    
    rGraph.Range("a1:a6").NumberFormat = "@"
    Set oChart = oWB.Charts.Add
    With oChart
    'chart title
        .HasTitle = True
        .ChartTitle.Characters.Text = oWS.Name
        
    'legend position
    .HasLegend = False
        '.Legend.Position = xlLegendPositionTop
    'chart type
    'ChartObjects.Add(50, 40, 300, 200).Chart

        .ChartType = xlLineMarkers
        
        .Axes(xlValue, xlPrimary).HasMajorGridlines = True
        .Axes(xlValue, xlPrimary).MajorGridlines.Border.LineStyle = xlDot
        .Axes(xlValue, xlPrimary).MajorGridlines.Border.ColorIndex = 15
        .SetSourceData Source:=rGraph, PlotBy:=xlColumns
        .Axes(xlCategory).HasTitle = True
       .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = "year"
        .Location Where:=xlLocationAsNewSheet, Name:="Chr"
        For i = 0 To UBound(arrEthnicities) - 1 'category white,.....
            rGraph(1, 2 + i).Value = oWS.Range("a" & arrEthnicities(i)).Value & " " & oWS.Range("b" & arrEthnicities(i)).Value
            For j = 0 To Year(Date) - 2004
                
                rGraph(j + 2, 2 + i).Value = oWS.Range("c" & (30 * j + (7 + i))).Value
                rGraph(j + 2, 2 + i).NumberFormat = "0.00"
            Next j
            .SeriesCollection(i + 1).XValues = rGraph.Range("a2:a6")
            .SeriesCollection(i + 1).Values = rGraph.Range(rGraph.Cells(2, 2 + i), rGraph.Cells(6, 2 + i))
            
            .SeriesCollection(i + 1).Points(.SeriesCollection(i + 1).Points.Count - 1).ApplyDataLabels AutoText:=True, _
        LegendKey:=False, ShowSeriesName:=True, ShowCategoryName:=False, _
        ShowValue:=False, ShowPercentage:=False, ShowBubbleSize:=False
            .SeriesCollection(i + 1).Name = rGraph(1, 2 + i).Value
            
        With .SeriesCollection(i + 1).Border
           .ColorIndex = 5    'q2
           .Weight = xlThick
            .LineStyle = xlContinuous
            
        End With
            With .SeriesCollection(i + 1)
                .MarkerBackgroundColorIndex = xlAutomatic
                .MarkerForegroundColorIndex = xlAutomatic
                .MarkerStyle = xlAutomatic
                .Smooth = False
                .MarkerSize = 5
                .Shadow = False
            End With
            .ChartArea.Font.Size = 12
            .PlotArea.Interior.ColorIndex = 2

            .PlotArea.Interior.PatternColor = 3
            '.PlotArea.Height = 200000
          
            'ActiveChart.SeriesCollection(1).Points(11).
    ' ows.range("a1").Characters.
    Next i

        
   
    End With
   ' oChart.HeightPercent = 80
End Sub

Sub getComparison()
Dim intX As Integer

For intX = 1 To oWB.Worksheets.Count
    cmbGraph.AddItem oWB.Worksheets(intX).Name
Next intX
cmbGraph.Text = cmbGraph.List(0)
Set oWS = oWB.Worksheets.Item(cmbGraph.Text)
End Sub

'**********************************Format Control ****************************************************
Private Sub formatControls()
    Dim i As Integer
    Me.Width = 3 * Screen.Width / 4
    Me.Height = 3 * Screen.Height / 4
    Me.Left = Screen.Width / 8
    Me.Top = Screen.Width / 16
    
    fraSource.Height = 5 * chckSource(0).Height
    fraSource.Width = (fraMain.Width - 1.5 * fraMain.Left) / 4
    fraMain.Width = 0.9 * Me.Width
    fraMain.Height = 0.78 * Me.Height
    fraMain.Top = (Me.Height - fraMain.Height) / 6
    fraMain.Left = (Me.Width - fraMain.Width) / 2
    
    fraChart.Width = (fraMain.Width - 1.5 * fraMain.Left - fraSource.Width) / 2
    fraChart.Left = fraMain.Left / 3
    fraChart.Height = fraSource.Height
    fraChart.Top = fraMain.Top
    
    fraCat.Width = fraChart.Width
    fraCat.Left = 2 * fraChart.Left + fraChart.Width
    fraCat.Height = fraChart.Height
    fraCat.Top = fraMain.Top
    
    fraSource.Left = fraCat.Left + fraCat.Width + fraChart.Left
    fraSource.Top = fraChart.Top

    cmbGraph.Width = 0.9 * fraChart.Width
    cmbGraph.Left = fraChart.Left
    cmbGraph.Top = fraChart.Top
    
    cmbField.Width = 0.9 * fraChart.Width
    cmbField.Left = fraChart.Left
    cmbField.Top = fraChart.Top + cmbField.Height
    
    cmbCat.Width = 0.95 * fraCat.Width
    cmbCat.Left = cmbGraph.Left / 2
    cmbCat.Top = (fraCat.Height - cmbCat.Height) / 2
    
  
    
    fraMCat.Top = fraChart.Height + fraChart.Top
    fraMCat.Height = fraMain.Height - fraMCat.Top - fraChart.Top
    fraMCat.Width = 3 * fraChart.Width / 4
    fraMCat.Left = fraChart.Left
    
    fraSec.Left = fraMCat.Left
    fraSec.Width = fraMCat.Width
    fraSec.Top = fraMCat.Top
    fraSec.Height = fraMCat.Height
    fraGraph.Left = 2 * fraMCat.Left + fraMCat.Width
    fraGraph.Top = fraMCat.Top
   
    fraGraph.Height = fraMCat.Height
    fraGraph.Width = (fraMain.Width - fraMCat.Width) - 3 * fraMCat.Left

    
    
    lstMain.Left = fraMCat.Left
    lstMain.Width = fraMCat.Width - 2 * lstMain.Left
    lstMain.Top = 0.1 * fraMCat.Height
    lstMain.Height = fraMCat.Height - 1.5 * lstMain.Top
    lstSecondry.Top = lstMain.Top
    lstSecondry.Height = lstMain.Height
    lstSecondry.Width = lstMain.Width
    lstSecondry.Left = lstMain.Left
    
    cmdExcel.Height = (Me.Height - (fraMain.Top + fraMain.Height)) / 4
    cmdExcel.Width = 0.15 * Me.Width
    cmdExcel.Left = (Me.Width - 3 * cmdExcel.Width) / 2
    
    cmdExcel.Top = fraMain.Top + fraMain.Height + cmdExcel.Height / 2
    cmdExit.Height = cmdExcel.Height
    cmdExit.Width = cmdExcel.Width
    cmdExit.Left = cmdExcel.Left + 1.5 * cmdExcel.Width
    cmdExit.Top = cmdExcel.Top
    imgLogo.Height = Me.Height - fraMain.Height - 3 * fraMain.Top
    imgLogo.Top = fraMain.Height + fraMain.Top
    imgLogo.Left = 0
    lblList.Top = lstMain.Top - lblList.Height
End Sub

Private Sub loadLists()
    lstSecondry.AddItem "White British"
    lstSecondry.AddItem "White Irish"
    lstSecondry.AddItem "White Other White"
    lstSecondry.AddItem "Mixed White and Black Caribbean"
    lstSecondry.AddItem "Mixed White and Black African"
    lstSecondry.AddItem "mixed White And Asian"
    lstSecondry.AddItem "Mixed Other mixed"
    lstSecondry.AddItem "Asian or Asian British Indian"
    lstSecondry.AddItem "Asian or Asian British Pakistani"
    lstSecondry.AddItem "Asian or Asian British Bangladeshi"
    lstSecondry.AddItem "Asian or Asian British Other Asian"
    lstSecondry.AddItem "Black or Black British Caribbean"
    lstSecondry.AddItem "Black or Black British African"
    lstSecondry.AddItem "Black or Black British Other Black"
    lstSecondry.AddItem "Other Ethnic Groups Chinese"
    lstSecondry.AddItem "Other Ethnic Groups Other"
    lstSecondry.AddItem "Not Stated  "



    lstMain.AddItem "White"
    lstMain.AddItem "Mixed"
    lstMain.AddItem "Asian or Asian British"
    lstMain.AddItem "Black or Black British"
    lstMain.AddItem "Other Ethnic Groups"
    lstMain.AddItem "Not Stated"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cleanup:
   Set oWS = Nothing
If Not oWB Is Nothing Then oWB.Close False
    Set oWB = Nothing
    oExcel.Quit
    Set oExcel = Nothing
End Sub

Private Sub lstSecondry_Click()
    Dim i As Integer
    If lstSecondry.SelCount = 6 Then
        MsgBox "you are allowed five categories only." & vbCrLf & "If you need to change the selected items click on them again"
        lstSecondry.Selected(lstSecondry.ListIndex) = False
    End If
    'reset
    'For i = 0 To 4
       'arrEthnicities(i) = 0
    'Next i
    
    'For i = 0 To lstSecondry.SelCount - 1
   
       ' arrEthnicities(i) = lstSecondry.Selected(i)
   ' Next i
    
End Sub
