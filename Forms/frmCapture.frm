VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmCapture 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "vbCPUID Capture"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Capture destination"
      Height          =   1020
      Left            =   150
      TabIndex        =   3
      Top             =   877
      Width           =   2745
      Begin VB.OptionButton optCapture 
         Caption         =   "Capture to clipboard"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   645
         Width           =   2460
      End
      Begin VB.OptionButton optCapture 
         Caption         =   "Capture to bitmap file"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   2460
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3075
      Top             =   2070
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1597
      TabIndex        =   1
      Top             =   2175
      Width           =   1110
   End
   Begin VB.CommandButton cmdCapture 
      Caption         =   "Capture"
      Default         =   -1  'True
      Height          =   345
      Left            =   337
      TabIndex        =   0
      Top             =   2175
      Width           =   1110
   End
   Begin VB.Label lblInstruction 
      AutoSize        =   -1  'True
      Caption         =   "Select the destination..."
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   390
      Width           =   2250
   End
   Begin VB.Label lblInstruction 
      AutoSize        =   -1  'True
      Caption         =   "Select the tab to capture"
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2355
   End
End
Attribute VB_Name = "frmCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
' frmCapture - Capture the VBCPUID app's screen to bitmap or clipboard
'
'==============================================================================
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub Form_Load()
  Set Icon = frmVBCPUID.Icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  If UnloadMode = vbFormControlMenu Then
    Cancel = True
    Call Hide
  End If
End Sub

Private Sub cmdCancel_Click()
  Call Hide
End Sub

Private Sub cmdCapture_Click()
  Const PATH_BITMAP As String = "PathBitmap"
  Dim sFile   As String
  
  If optCapture(0).Value Then
    With cd
      .CancelError = True
      .DefaultExt = "bmp"
      .DialogTitle = "Save as bitmap"
      .Filter = "Bitmaps (*.bmp)|*.bmp"
      .FilterIndex = 1
      .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
      .InitDir = GetSetting(App.ProductName, PATH_BITMAP, PATH_BITMAP, CurDir$)
      .MaxFileSize = 255
      
      On Error GoTo Catch
      .ShowSave
      
      sFile = .FileName
      Call SaveSetting(App.ProductName, PATH_BITMAP, PATH_BITMAP, frmVBCPUID.GetPath(sFile))
    End With
  End If
  
  Call frmVBCPUID.SetFocus
  DoEvents
  
  Call Clipboard.Clear
  Call keybd_event(vbKeySnapshot, 1, 0, 0)
  DoEvents
  
  If optCapture(0).Value Then
    Call SavePicture(Clipboard.GetData, sFile)
    Call Clipboard.Clear
  End If
  
  Call SetFocus
Catch:
  On Error GoTo 0
End Sub
