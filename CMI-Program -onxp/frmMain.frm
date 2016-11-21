VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "CMI Analyser"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame fraMain 
      Height          =   4215
      Left            =   2160
      TabIndex        =   1
      Top             =   960
      Width           =   7575
      Begin VB.CommandButton cmdLocal 
         Caption         =   "Process Raw Local Data"
         Height          =   735
         Left            =   3000
         TabIndex        =   6
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate Aggregate Data"
         Height          =   735
         Left            =   600
         TabIndex        =   5
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton cmdData 
         Caption         =   "Add/Change/Delete Input Excel Files Locations"
         Height          =   735
         Left            =   600
         TabIndex        =   3
         Top             =   720
         Width           =   2295
      End
      Begin VB.CommandButton cmdAnalysis 
         Caption         =   "Analysis and Graphical Representation"
         Height          =   735
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   3360
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Image imgLogo 
      Height          =   1740
      Left            =   0
      Picture         =   "frmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   4440
      Width           =   1635
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnalysis_Click()
  
    lblMsg.Caption = " Please wait while data is being analysed... this might take a few minutes"
    Me.FontSize = 14
    Me.Caption = "!!!!!! Please Wait....."
    hideControls

    frmGraph.Show
    Unload Me
End Sub


Private Sub cmdData_Click()
    frmSelect.Show
    Unload Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    generateData
    MsgBox "Data analysis is now complete, Click ok to continue"
End Sub

Private Sub cmdLocal_Click()
    Dim strFile As String
    Dim i As Integer
    Dim wsName() As String
    strFile = "C:\Users\User\Documents\CMI\Count me in project\Count Me in\National\Count me in census 2005 England results.xls"

End Sub

Private Sub Form_Load()
   
    formatControls
 
End Sub


Private Sub formatControls()
    Dim i As Integer
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.Left = Screen.Width
    Me.Top = Screen.Width
    
    fraMain.Width = 0.8 * Me.Width
    fraMain.Height = 0.6 * Me.Height
    fraMain.Top = (Me.Height - fraMain.Height) / 4
    fraMain.Left = (Me.Width - fraMain.Width) / 2
    
    cmdData.Width = (fraMain.Width - 1.5 * fraMain.Left) / 2
    cmdData.Left = fraMain.Left / 2
    cmdData.Height = fraMain.Height / 7
    cmdData.Top = fraMain.Top
    
    cmdAnalysis.Width = cmdData.Width
    cmdAnalysis.Left = fraMain.Left + cmdData.Width
    cmdAnalysis.Height = cmdData.Height
    cmdAnalysis.Top = fraMain.Top
    cmdExit.Left = (Me.Width - cmdExit.Width) / 2
    cmdExit.Top = (fraMain.Top + fraMain.Height) + (Me.Height - (fraMain.Top + fraMain.Height) - cmdExit.Height) / 2
    
    With cmdGenerate
        .Width = cmdData.Width
        .Height = cmdData.Height
        .Left = cmdData.Left
        .Top = cmdData.Top + 2 * cmdData.Height
    End With
    
    With cmdLocal
        .Width = cmdAnalysis.Width
        .Left = cmdAnalysis.Left
        .Height = cmdAnalysis.Height
        .Top = cmdGenerate.Top
        
        
    End With
    With lblMsg
        .WordWrap = True
        .Left = fraMain.Width / 5
        .Width = fraMain.Width - 2 * .Left
        
        .Top = fraMain.Top + (fraMain.Height - .Height) / 4
    End With
    With lblMsg.Font
        .Size = 14
        .Bold = True
    End With
    lblMsg.ForeColor = vbRed
    lblMsg.Font = "Arial"
    With imgLogo
    .Height = Me.Height / 8
    .Width = Me.Width / 10
    .Left = 0
    .Top = 0
    End With
End Sub


Private Sub hideControls()
    cmdData.Visible = False
    cmdAnalysis.Visible = False
    cmdExit.Visible = False
    cmdGenerate.Visible = False
    cmdLocal.Visible = False
End Sub

