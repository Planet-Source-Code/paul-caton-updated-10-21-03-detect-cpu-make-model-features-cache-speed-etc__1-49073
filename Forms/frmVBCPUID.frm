VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmVBCPUID 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
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
   Icon            =   "frmVBCPUID.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9735
   Begin VB.PictureBox picTab 
      Height          =   3540
      Index           =   5
      Left            =   3270
      ScaleHeight     =   3480
      ScaleWidth      =   6180
      TabIndex        =   6
      Top             =   3285
      Visible         =   0   'False
      Width           =   6240
      Begin VB.PictureBox picApp 
         Height          =   1020
         Left            =   2227
         ScaleHeight     =   960
         ScaleWidth      =   1680
         TabIndex        =   36
         Top             =   2265
         Width           =   1740
         Begin VB.Label lblVersion 
            Alignment       =   2  'Center
            Caption         =   "Version"
            Height          =   210
            Left            =   60
            TabIndex        =   37
            Top             =   690
            Width           =   1575
         End
         Begin VB.Image imgApp 
            Appearance      =   0  'Flat
            Height          =   480
            Left            =   600
            Picture         =   "frmVBCPUID.frx":27A2
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   1020
         Left            =   4215
         Picture         =   "frmVBCPUID.frx":2D40
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   " Exit VBCPUID "
         Top             =   2265
         Width           =   1740
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Read CPU"
         Height          =   1020
         Left            =   240
         Picture         =   "frmVBCPUID.frx":2E4E
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   " Re-read the CPU "
         Top             =   2265
         Width           =   1740
      End
      Begin VB.TextBox txtCaption 
         Height          =   315
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   5715
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save Data..."
         Height          =   1020
         Left            =   2227
         Picture         =   "frmVBCPUID.frx":32AD
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   " Save the VBCPUID data "
         Top             =   1020
         Width           =   1740
      End
      Begin VB.CommandButton cmdLoad 
         Caption         =   "&Load Data..."
         Height          =   1020
         Left            =   240
         Picture         =   "frmVBCPUID.frx":372D
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   " Load a VBCPUID data file "
         Top             =   1020
         Width           =   1740
      End
      Begin VB.CommandButton cmdCapture 
         Caption         =   "&Capture..."
         Height          =   1020
         Left            =   4215
         Picture         =   "frmVBCPUID.frx":3B73
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   " Capture the VBCPUID image to bitmap or the clipboard "
         Top             =   1020
         Width           =   1740
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   5700
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblCaption 
         AutoSize        =   -1  'True
         Caption         =   "Caption:"
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   810
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   3540
      Index           =   4
      Left            =   2664
      ScaleHeight     =   3480
      ScaleWidth      =   6180
      TabIndex        =   5
      Top             =   2745
      Visible         =   0   'False
      Width           =   6240
      Begin MSComctlLib.ListView lvReg 
         Height          =   3255
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Level"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "eax"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "ebx"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "ecx"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "edx"
            Object.Width           =   2064
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   3540
      Index           =   3
      Left            =   2058
      ScaleHeight     =   3480
      ScaleWidth      =   6180
      TabIndex        =   16
      Top             =   2205
      Visible         =   0   'False
      Width           =   6240
      Begin VB.Frame fmPower 
         Caption         =   "Power management"
         Enabled         =   0   'False
         Height          =   3060
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox chkPm 
            Caption         =   "Software thermal control"
            Height          =   210
            Index           =   5
            Left            =   225
            TabIndex        =   27
            Top             =   2610
            Width           =   2685
         End
         Begin VB.CheckBox chkPm 
            Caption         =   "Thermal monitoring"
            Height          =   210
            Index           =   4
            Left            =   225
            TabIndex        =   26
            Top             =   2190
            Width           =   2685
         End
         Begin VB.CheckBox chkPm 
            Caption         =   "Thermal trip"
            Height          =   210
            Index           =   3
            Left            =   225
            TabIndex        =   25
            Top             =   1770
            Width           =   2685
         End
         Begin VB.CheckBox chkPm 
            Caption         =   "Voltage ID control"
            Height          =   210
            Index           =   2
            Left            =   225
            TabIndex        =   24
            Top             =   1350
            Width           =   2685
         End
         Begin VB.CheckBox chkPm 
            Caption         =   "Frequency ID control"
            Height          =   210
            Index           =   1
            Left            =   225
            TabIndex        =   23
            Top             =   930
            Width           =   2685
         End
         Begin VB.CheckBox chkPm 
            Caption         =   "Temperature sensor"
            Height          =   210
            Index           =   0
            Left            =   225
            TabIndex        =   22
            Top             =   510
            Width           =   2685
         End
      End
      Begin VB.Frame frAddress 
         Caption         =   "Address bits"
         Enabled         =   0   'False
         Height          =   1380
         Left            =   3870
         TabIndex        =   18
         Top             =   240
         Width           =   2085
         Begin VB.Label lblAddrVirt 
            Height          =   210
            Left            =   1125
            TabIndex        =   29
            Top             =   930
            Width           =   240
         End
         Begin VB.Label lblAddrPhys 
            Height          =   210
            Left            =   1125
            TabIndex        =   28
            Top             =   510
            Width           =   240
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            Caption         =   "Virtual:"
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   930
            Width           =   675
         End
         Begin VB.Label lblAddress 
            AutoSize        =   -1  'True
            Caption         =   "Physical:"
            Height          =   210
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   510
            Width           =   825
         End
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   3540
      Index           =   2
      Left            =   1452
      ScaleHeight     =   3480
      ScaleWidth      =   6180
      TabIndex        =   4
      Top             =   1665
      Visible         =   0   'False
      Width           =   6240
      Begin MSComctlLib.ListView lvCache 
         Height          =   3240
         Left            =   120
         TabIndex        =   17
         Top             =   135
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   5715
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Assoc"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Entries"
            Object.Width           =   2064
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Size KB"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Line bytes"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   3540
      Index           =   1
      Left            =   846
      ScaleHeight     =   3480
      ScaleWidth      =   6180
      TabIndex        =   1
      Top             =   1125
      Visible         =   0   'False
      Width           =   6240
      Begin MSComctlLib.ListView lvFeatures 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   5741
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   1799
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   7990
         EndProperty
      End
   End
   Begin VB.PictureBox picTab 
      Height          =   3555
      Index           =   0
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   6180
      TabIndex        =   2
      Top             =   570
      Width           =   6240
      Begin MSComctlLib.ListView lvProc 
         Height          =   2820
         Left            =   90
         TabIndex        =   3
         Top             =   120
         Width           =   5970
         _ExtentX        =   10530
         _ExtentY        =   4974
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         PictureAlignment=   1
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Property"
            Object.Width           =   2937
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Value"
            Object.Width           =   7444
         EndProperty
      End
      Begin VB.Frame fm 
         Height          =   450
         Left            =   120
         TabIndex        =   9
         Top             =   2955
         Width           =   5955
         Begin VB.Label lblSpeedFull 
            AutoSize        =   -1  'True
            Caption         =   "Speed:"
            Height          =   210
            Left            =   90
            TabIndex        =   15
            ToolTipText     =   " Measured cpu clock speed "
            Top             =   165
            Width           =   675
         End
         Begin VB.Label lblMHzFull 
            Caption         =   "8,888 MHz"
            Height          =   210
            Left            =   810
            TabIndex        =   14
            ToolTipText     =   " Measured cpu clock speed "
            Top             =   165
            Width           =   1020
         End
         Begin VB.Label lblMHzCurr 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0,000.000 MHz"
            Height          =   210
            Left            =   4425
            TabIndex        =   13
            ToolTipText     =   "Real time cpu clock speed "
            Top             =   165
            Width           =   1410
         End
         Begin VB.Label lblSpeedCurr 
            AutoSize        =   -1  'True
            Caption         =   "Real Time Speed:"
            Height          =   210
            Left            =   2745
            TabIndex        =   12
            ToolTipText     =   "Real time cpu clock speed "
            Top             =   165
            Width           =   1635
         End
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   11
      Top             =   7020
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   106
            Text            =   "MMX"
            TextSave        =   "MMX"
            Object.ToolTipText     =   " Intel multimedia extensions "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   847
            MinWidth        =   106
            Text            =   "SSE"
            TextSave        =   "SSE"
            Object.ToolTipText     =   " Intel streaming simd extensions "
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   106
            Text            =   "SSE2"
            TextSave        =   "SSE2"
            Object.ToolTipText     =   " Intel streaming simd extensions 2 "
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   106
            Text            =   "SSE3"
            TextSave        =   "SSE3"
            Object.ToolTipText     =   " Intel streaming simd extensions 3 "
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1191
            MinWidth        =   106
            Text            =   "MMX+"
            TextSave        =   "MMX+"
            Object.ToolTipText     =   " Cyrix/AMD multimedia extensions "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1535
            MinWidth        =   106
            Text            =   "3DNow!"
            TextSave        =   "3DNow!"
            Object.ToolTipText     =   " AMD multimedia extensions "
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   106
            Text            =   "3DNow!+"
            TextSave        =   "3DNow!+"
            Object.ToolTipText     =   " AMD multimedia extensions + "
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   8520
            MinWidth        =   106
            Text            =   "Hyper-Threading"
            TextSave        =   "Hyper-Threading"
            Object.ToolTipText     =   " Intel Hyper-Threading technology "
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7005
      Top             =   165
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   4140
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   7303
      TabWidthStyle   =   1
      TabFixedWidth   =   1879
      HotTracking     =   -1  'True
      Separators      =   -1  'True
      TabMinWidth     =   0
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Processor"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Features"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cache"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "AMD Extra"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Registers"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Actions"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmVBCPUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'VBCPUID - A CPUID (Central Processing Unit IDentity) type application to
'demonstrate the cCpuInfo class
'
'If you see anything unusual with your cpu's reported identity, feel free to
'save the data (see 'Actions' tab) and mail the file to me...
'Paul_Caton@hotmail.com
'
'031007 v1.00 First cut........................................................
'031008 v1.01 CPUID extended levels up to 80000004.............................
'031008 v1.02 Added screen capture, conditionalised display of S/N.............
'031013 v1.03 Abstracted data decode to the cCpuInfo class
'             Decode for all processors........................................
'031014 v1.04 Cache decode for AMD
'             Load/Save data
'             Extra tab for AMD
'             CPUID extended levels up to 80000008.............................
'031015 v1.05 Tidy up and few minor fixes......................................
'031020 v1.06 Fixed problem when compiled native and optimized.................
'
'Copyright free, use as and how you please.
'==============================================================================
Option Explicit

'Data file constants
Private Const PATH_DATA                 As String = "PathText"
Private Const SECTION_APP               As String = "App"
Private Const SECTION_SPEED             As String = "Speed"
Private Const SECTION_LEVELS            As String = "Levels"
Private Const KEY_NAME                  As String = "Name"
Private Const KEY_VERSION               As String = "Version"
Private Const KEY_CAPTION               As String = "Caption"
Private Const KEY_DATE                  As String = "Date"
Private Const KEY_TIME                  As String = "Time"
Private Const KEY_SPEED                 As String = "Speed"
Private Const KEY_LEVEL                 As String = "Level_"

'ListView column sizing constants
Private Const LVM_FIRST                 As Long = &H1000
Private Const LVM_SETCOLUMNWIDTH        As Long = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE            As Long = -1
Private Const LVSCW_AUTOSIZE_USEHEADER  As Long = -2
    
'Api declaration
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'xCpuInfo class instance
Private cpu       As cCpuInfo
Private sLoadFile As String
Private itm       As MSComctlLib.ListItem

Private Sub Form_Load()
  Dim i       As Long
  Dim nHeight As Long
  Dim nWidth  As Long
  
'Load the 'Capture' form
  Call Load(frmCapture)
  
'Position the tab control
  Call ts.Move(120, 120, 6495, 4140)
  
'Position the tab pictures
  For i = 0 To 5
    Call picTab(i).Move(240, 570, 6240, 3555)
  Next i
  
'Position the listview controls
  Call lvProc.Move(120, 135, 5955, 2820)
  Call lvFeatures.Move(120, 135, 5955, 3255)
  Call lvCache.Move(120, 135, 5955, 3255)
  Call lvReg.Move(120, 135, 5955, 3255)
  
'Position the form
  nWidth = (Width - ScaleWidth) + ts.Left + ts.Width + 120
  nHeight = (Height - ScaleHeight) + ts.Top + ts.Height + 120 + sb.Height
  Call Move((Screen.Width - nWidth) / 2, (Screen.Height - nHeight) / 4, nWidth, nHeight)

'Add the feature items
  With lvFeatures.ListItems
    Set itm = .Add(, , " FPU"):    itm.SubItems(1) = "x87 FPU on Chip"
    Set itm = .Add(, , " VME"):    itm.SubItems(1) = "Virtual 8086 Mode Enhancement"
    Set itm = .Add(, , " DE"):     itm.SubItems(1) = "Debugging Extensions"
    Set itm = .Add(, , " PSE"):    itm.SubItems(1) = "Page Size Extensions"
    Set itm = .Add(, , " TSC"):    itm.SubItems(1) = "Time Stamp Counter"
    Set itm = .Add(, , " MSR"):    itm.SubItems(1) = "RDMSR and WRMSR Support"
    Set itm = .Add(, , " PAE"):    itm.SubItems(1) = "Physical Address Extensions"
    Set itm = .Add(, , " MCE"):    itm.SubItems(1) = "Machine Check Exception"
    Set itm = .Add(, , " CX8"):    itm.SubItems(1) = "CMPXCHG8B instruction"
    Set itm = .Add(, , " APIC"):   itm.SubItems(1) = "APIC on Chip"
    Set itm = .Add(, , " SEP"):    itm.SubItems(1) = "SYSENTER SYSEXIT"
    Set itm = .Add(, , " MTRR"):   itm.SubItems(1) = "Memory Type Range Registers"
    Set itm = .Add(, , " PGE"):    itm.SubItems(1) = "PTE Global Bit"
    Set itm = .Add(, , " MCA"):    itm.SubItems(1) = "Machine Check Architecture"
    Set itm = .Add(, , " CMOV"):   itm.SubItems(1) = "Conditional Move/Compare instuction"
    Set itm = .Add(, , " PAT"):    itm.SubItems(1) = "Page Attribute Table"
    Set itm = .Add(, , " PSE36"):  itm.SubItems(1) = "Page Size Extension"
    Set itm = .Add(, , " PSN"):    itm.SubItems(1) = "Processor Serial Number"
    Set itm = .Add(, , " CLFSH"):  itm.SubItems(1) = "CFLUSH instruction"
    Set itm = .Add(, , " DS"):     itm.SubItems(1) = "Debug Store"
    Set itm = .Add(, , " ACPI"):   itm.SubItems(1) = "Thermal Monitor and Clock control"
    Set itm = .Add(, , " FXSR"):   itm.SubItems(1) = "FXSAVE/FXRSTOR instructions"
    Set itm = .Add(, , " SS"):     itm.SubItems(1) = "Self Snoop"
    Set itm = .Add(, , " HTT"):    itm.SubItems(1) = "Hyper-Threading Technology"
    Set itm = .Add(, , " TM1"):    itm.SubItems(1) = "Thermal Monitor"
    Set itm = .Add(, , " IA64"):   itm.SubItems(1) = "IA-64 jump instructions"
    Set itm = .Add(, , " PBE"):    itm.SubItems(1) = "Pending Break Enable"
    Set itm = .Add(, , " EST"):    itm.SubItems(1) = "Enhanced SpeedStep Technology"
    Set itm = .Add(, , " TM2"):    itm.SubItems(1) = "Thermal Monitor 2 Technology"
    Set itm = .Add(, , " CID"):    itm.SubItems(1) = "Context ID"
  End With
  Call AutoSize(lvFeatures)

'Create a cCpuInfo instance
  Set cpu = New cCpuInfo
  Call ProcessData
  
'Set the caption text box (sets the caption also via txtCaption_Change)
  txtCaption.Text = App.Title
  lblVersion.Caption = "v" & FmtVersion
  
  Call CalcSpeed
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Call Unload(frmCapture)
  Set cpu = Nothing
End Sub

'Capture the app's image to the clipboard or a bitmap file
Private Sub cmdCapture_Click()
  Call frmCapture.Show(vbModeless)
End Sub

Private Sub cmdExit_Click()
  Call Unload(Me)
End Sub

'Load a VBCPUID data file
Private Sub cmdLoad_Click()
  Dim i     As Long
  Dim eax   As Long
  Dim ebx   As Long
  Dim ecx   As Long
  Dim edx   As Long
  Dim sFile As String
  Dim s()   As String
  Dim ini   As cIni
  
'Show the file selector dialog
  With cd
    .CancelError = True
    .DefaultExt = "txt"
    .DialogTitle = "Load data"
    .Filter = "Text files (*.txt)|*.txt"
    .FilterIndex = 1
    .Flags = cdlOFNHideReadOnly
    
    'Remember the last directory used
    .InitDir = GetSetting(App.Title, PATH_DATA, PATH_DATA, CurDir$)
    .MaxFileSize = 255
    
    On Error GoTo CatchCancel
    Call .ShowOpen
    
    sFile = .FileName
    sLoadFile = GetFile(sFile)
    
    'Save the directory used
    Call SaveSetting(App.Title, PATH_DATA, PATH_DATA, GetPath(sFile))
  End With
  
  On Error GoTo CatchError
  
  'Create ini file instance
  Set ini = New cIni
  With ini
    .Path = sFile
    .Section = SECTION_APP
    .Key = KEY_NAME
    
    'Validate the data file
    If StrComp(.Value, App.Title, vbTextCompare) <> 0 Then
      MsgBox "Invalid " & App.Title & " data file", vbCritical, "Load data"
      GoTo CatchCancel
    End If
    
    'Load the data file's saved caption
    .Key = KEY_CAPTION
    txtCaption.Text = .Value
    Caption = txtCaption.Text & " [" & sLoadFile & "]"
    
    'Load the data files saved speed
    .Section = SECTION_SPEED
    .Key = KEY_SPEED
    lblMHzFull.Caption = .Value
    
    'Load the register values from the data file
    .Section = SECTION_LEVELS
    For i = eLevels.el_Std0 To eLevels.el_Xtd8
      
      .Key = KEY_LEVEL & Hex$(i)
      s = Split(.Value, " ")
      eax = Val("&H" & s(0))
      ebx = Val("&H" & s(1))
      ecx = Val("&H" & s(2))
      edx = Val("&H" & s(3))
      
      'Set the register values in the cCpuInfo class
      Call cpu.LevelSet(i, eax, ebx, ecx, edx)
    Next i
    
    Call cpu.Refresh(False)             'Process the new data in the cCpuInfo class. The False parameter means use the registers as they've been set, don't re-read from the cpu
    Call ProcessData                    'Decode and display the data
    tmrSpeed.Enabled = False            'Disable the real time speed timer
    lblMHzCurr.Caption = vbNullString   'Blank the real time speed
  End With

'Catch the file dialog cancel
CatchCancel:
  Set ini = Nothing
  On Error GoTo 0
  Exit Sub

'Catch error
CatchError:
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdLoad_Click of Form frmVBCPUID"
  Resume CatchCancel
End Sub

'Re-read the cpu registers... do this after loading a data file to return to the running cpu
Private Sub cmdRefresh_Click()
  Call cpu.Refresh(True)            'Re-read the cpu registers
  Call ProcessData                  'Display the data
  sLoadFile = vbNullString
  Caption = App.Title               'Reset the caption
  txtCaption.Text = Caption
  Call CalcSpeed
End Sub

'Save the VBCPUID data
Private Sub cmdSave_Click()
  Dim i     As Long
  Dim eax   As Long
  Dim ebx   As Long
  Dim ecx   As Long
  Dim edx   As Long
  Dim sFile As String
  Dim ini   As cIni
  
  'Text file class
  Set ini = New cIni
  
  'File save dialog
  With cd
    .CancelError = True
    .DefaultExt = "txt"
    .DialogTitle = "Save data"
    .Filter = "Text files (*.txt)|*.txt"
    .FilterIndex = 1
    .Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist + cdlOFNNoReadOnlyReturn + cdlOFNHideReadOnly
    
    'Remember the last directory used
    .InitDir = GetSetting(App.Title, PATH_DATA, PATH_DATA, CurDir$)
    .MaxFileSize = 255
    
    'Catch Cancel
    On Error GoTo Catch
    .ShowSave
    
    sFile = .FileName
    
    'Save the directory used
    Call SaveSetting(App.Title, PATH_DATA, PATH_DATA, GetPath(sFile))
  End With
  
  ini.Path = sFile
  
  With ini
    .Section = SECTION_APP
    .Key = KEY_NAME
    .Value = App.Title
    
    .Key = KEY_CAPTION
    .Value = txtCaption.Text
    
    .Key = KEY_VERSION
    .Value = FmtVersion
    
    .Key = KEY_DATE
    .Value = Format$(Date$, "Long Date")
    
    .Key = KEY_TIME
    .Value = Format$(Time$, "Long Time")
    
    .Section = SECTION_SPEED
    .Key = KEY_SPEED
    .Value = lblMHzFull.Caption
    
    .Section = SECTION_LEVELS
    For i = eLevels.el_Std0 To eLevels.el_Xtd8
      Call cpu.LevelGet(i, eax, ebx, ecx, edx)
      .Key = KEY_LEVEL & Hex$(i)
      .Value = cpu.HexPad(eax) & " " & cpu.HexPad(ebx) & " " & cpu.HexPad(ecx) & " " & cpu.HexPad(edx)
    Next i
  End With
  
Catch:
  On Error GoTo 0
  Set ini = Nothing
End Sub

Private Sub lvFeatures_ItemCheck(ByVal Item As MSComctlLib.ListItem)
  'Checkbox clicked, restore it's state
  Item.Checked = Not Item.Checked
End Sub

'Real time speed timer
Private Sub tmrSpeed_Timer()
  Dim fSecs     As Double
  Dim cCycles   As Currency

  'Stop the benchmark
  Call cpu.BenchStop(fSecs, cCycles)
  
  'Display the real time speed
  lblMHzCurr.Caption = Format$(Round((cCycles / fSecs) / 1000000, 6), "#,###.000") & " MHz"
  
  'Start the benchmark for the next update
  cpu.BenchStart
End Sub

'Tab sheet click
Private Sub ts_Click()
  Static nTab As Long
  
  picTab(nTab).Visible = False
  nTab = ts.SelectedItem.Index - 1
  picTab(nTab).Visible = True
  
  Select Case nTab
  Case 5
    txtCaption.SetFocus
    Call SendKeys("{END}")
  End Select
End Sub

'Update the caption
Private Sub txtCaption_Change()
  If Len(sLoadFile) = 0 Then
    Caption = txtCaption.Text
  Else
    Caption = txtCaption.Text & " [" & sLoadFile & "]"
  End If
End Sub

'Process and display the cCpuInfo data
Private Sub ProcessData()
  Const GREY    As Long = &H808080
  Dim i         As Long
  Dim eax       As Long
  Dim ebx       As Long
  Dim ecx       As Long
  Dim edx       As Long
  Dim nLast     As Long
  Dim nColor    As Long
  Dim bArray()  As Byte
  Dim s         As String
  Dim sAssoc    As String
  Dim sEntries  As String
  Dim sSize     As String
  Dim sLine     As String
  Dim sSector   As String
  
'Processor tab
  Call lvProc.ListItems.Clear
  
  Set itm = lvProc.ListItems.Add(, , "Vendor ID")
  itm.SubItems(1) = cpu.VendorIdStr

  Set itm = lvProc.ListItems.Add(, , "Manufacturer")
  itm.SubItems(1) = cpu.ManufacturerStr
  
  s = cpu.NameStr
  If Len(s) > 0 Then
    Set itm = lvProc.ListItems.Add(, , "CPU Name")
    itm.SubItems(1) = s
  End If

  Set itm = lvProc.ListItems.Add(, , "CPU Type")
  itm.SubItems(1) = cpu.CpuTypeStr

  Set itm = lvProc.ListItems.Add(, , "CPU Family")
  itm.SubItems(1) = cpu.CpuFamilyStr

  Set itm = lvProc.ListItems.Add(, , "CPU Model")
  itm.SubItems(1) = cpu.CpuModelStr

  s = cpu.CpuBrandStr
  If Len(s) > 0 Then
    Set itm = lvProc.ListItems.Add(, , "CPU Brand")
    itm.SubItems(1) = s
  End If

  Set itm = lvProc.ListItems.Add(, , "CPU Stepping")
  itm.SubItems(1) = cpu.CpuStepping

  i = cpu.CpuLogicalCount
  If i > 0 Then
    Set itm = lvProc.ListItems.Add(, , "Logical CPU's")
    itm.SubItems(1) = i
  End If

  i = cpu.ApicID
  If i > 0 Then
    Set itm = lvProc.ListItems.Add(, , "APIC ID")
    itm.SubItems(1) = Hex$(i)
  End If

  Set itm = lvProc.ListItems.Add(, , "Serial Number")
  Select Case cpu.SnStatus
  Case eSnStatus.esn_Enabled:     itm.SubItems(1) = cpu.SnStr
  Case eSnStatus.esn_Disabled:    itm.SubItems(1) = "Disabled"
  Case eSnStatus.esn_Unavailable: itm.SubItems(1) = "Unavailable"
  End Select
  
  Call AutoSize(lvProc)

'Status bar
  With sb
    .Panels(1).Enabled = cpu.Feature(ef_MMX)
    .Panels(2).Enabled = cpu.Feature(ef_SSE)
    .Panels(3).Enabled = cpu.Feature(ef_SSE2)
    .Panels(4).Enabled = cpu.Feature(ef_SSE3)
    .Panels(5).Enabled = cpu.Feature(ef_x3DNOW)
    .Panels(6).Enabled = cpu.Feature(ef_x3DNOW_X)
    .Panels(7).Enabled = IIf(cpu.Vendor = ev_Cyrix, cpu.Feature(ef_xMMX_PLUS_CYRIX), cpu.Feature(ef_xMMX_PLUS))
    .Panels(8).Enabled = (cpu.Feature(ef_HTT) And cpu.CpuLogicalCount > 1)
  End With
  
'Features tab
  With lvFeatures
    Set itm = .ListItems(1):  itm.Checked = cpu.Feature(ef_FPU):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(2):  itm.Checked = cpu.Feature(ef_VME):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(3):  itm.Checked = cpu.Feature(ef_DE):    itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(4):  itm.Checked = cpu.Feature(ef_PSE):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(5):  itm.Checked = cpu.Feature(ef_TSC):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(6):  itm.Checked = cpu.Feature(ef_MSR):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(7):  itm.Checked = cpu.Feature(ef_PAE):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(8):  itm.Checked = cpu.Feature(ef_MCE):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(9):  itm.Checked = cpu.Feature(ef_CX8):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(10): itm.Checked = cpu.Feature(ef_APIC):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(11): itm.Checked = cpu.Feature(ef_SEP):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(12): itm.Checked = cpu.Feature(ef_MTRR):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(13): itm.Checked = cpu.Feature(ef_PGE):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(14): itm.Checked = cpu.Feature(ef_MCA):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(15): itm.Checked = cpu.Feature(ef_CMOV):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(16): itm.Checked = cpu.Feature(ef_PAT):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(17): itm.Checked = cpu.Feature(ef_PSE36): itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(18): itm.Checked = cpu.Feature(ef_PSN):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(19): itm.Checked = cpu.Feature(ef_CLFSH): itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(20): itm.Checked = cpu.Feature(ef_DS):    itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(21): itm.Checked = cpu.Feature(ef_ACPI):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(22): itm.Checked = cpu.Feature(ef_FXSR):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(23): itm.Checked = cpu.Feature(ef_SS):    itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(24): itm.Checked = cpu.Feature(ef_HTT):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(25): itm.Checked = cpu.Feature(ef_TM1):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(26): itm.Checked = cpu.Feature(ef_IA64):  itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(27): itm.Checked = cpu.Feature(ef_PBE):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(28): itm.Checked = cpu.Feature(ef_EST):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(29): itm.Checked = cpu.Feature(ef_TM2):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
    Set itm = .ListItems(30): itm.Checked = cpu.Feature(ef_CID):   itm.ForeColor = IIf(itm.Checked, 0, GREY)
  End With
  Call AutoSize(lvFeatures)
  
'TLB/Cache tab
  Call lvCache.ListItems.Clear
  lvCache.Sorted = False
  
  If cpu.LevelsExt > 4 Then
    'AMD extended cache descriptors
    lvCache.ColumnHeaders(6).Text = "Lines/tag"
    
    Set itm = lvCache.ListItems.Add(, , "L1 4K code pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_4KbCodeAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L1_4KbCodeEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L1 4K data pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_4KbDataAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L1_4KbDataEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L1 4M code pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_4MbCodeAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L1_4MbCodeEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L1 4M data pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_4MbDataAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L1_4MbDataEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L1 code cache")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_CodeAssoc))
    itm.SubItems(3) = cpu.xCache(ecf_L1_CodeSizeKb)
    itm.SubItems(4) = cpu.xCache(ecf_L1_CodeLineSizeBytes)
    itm.SubItems(5) = cpu.xCache(ecf_L1_CodeLinesPerTag)
    
    Set itm = lvCache.ListItems.Add(, , "L1 data cache")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L1_DataAssoc))
    itm.SubItems(3) = cpu.xCache(ecf_L1_DataSizeKb)
    itm.SubItems(4) = cpu.xCache(ecf_L1_DataLineSizeBytes)
    itm.SubItems(5) = cpu.xCache(ecf_L1_DataLinesPerTag)
    
    Set itm = lvCache.ListItems.Add(, , "L2 4K code pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L2_4KbCodeAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L2_4KbCodeEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L2 4K data pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L2_4KbDataAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L2_4KbDataEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L2 4M code pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L2_4MbCodeAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L2_4MbCodeEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L2 4M data pages")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L2_4MbDataAssoc))
    itm.SubItems(2) = cpu.xCache(ecf_L2_4MbDataEntries)
    
    Set itm = lvCache.ListItems.Add(, , "L2 unified cache")
    itm.SubItems(1) = cpu.xCacheAssocStr(cpu.xCache(ecf_L2_CacheAssoc))
    itm.SubItems(3) = cpu.xCache(ecf_L2_CacheSizeKb)
    itm.SubItems(4) = cpu.xCache(ecf_L2_CacheLineSizeBytes)
    itm.SubItems(5) = cpu.xCache(ecf_L2_CacheLinesPerTag)
  Else
    'Intel std cache descriptors
    bArray = cpu.CacheDescriptors

    lvCache.ColumnHeaders(6).Text = "Sector"
    
    For i = LBound(bArray) To UBound(bArray)
    
      s = vbNullString
      sAssoc = vbNullString
      sEntries = vbNullString
      sSize = vbNullString
      sLine = vbNullString
      sSector = vbNullString
      
      Call StdCacheDescriptors(bArray(i), s, sAssoc, sEntries, sSize, sLine, sSector)
      
      If Len(s) > 0 Then
        Set itm = lvCache.ListItems.Add(, , s)
        With itm
          .SubItems(1) = sAssoc
          .SubItems(2) = sEntries
          .SubItems(3) = sSize
          .SubItems(4) = sLine
          .SubItems(5) = sSector
        End With
      End If
    Next i
    
    lvCache.Sorted = True
  End If
  Call AutoSize(lvCache)
  
'AMD extra
  If cpu.LevelsExt > 4 Then
    lblAddrPhys.Caption = cpu.xAddrBitsPhysical
    lblAddrVirt.Caption = cpu.xAddrBitsVirtual
    
    For i = eExPowerManagment.epm_TemperatureSensor To eExPowerManagment.epm_SoftwareThermalControl
      chkPm(i).Value = IIf(cpu.xPowerManagement(i), vbChecked, vbUnchecked)
    Next i
  Else
    lblAddrPhys.Caption = ""
    lblAddrVirt.Caption = ""
    
    For i = eExPowerManagment.epm_TemperatureSensor To eExPowerManagment.epm_SoftwareThermalControl
      chkPm(i).Value = vbUnchecked
    Next i
  End If
  
'Registers tab
  Call lvReg.ListItems.Clear
  
  For i = 0 To 3
    Call cpu.LevelGet(i, eax, ebx, ecx, edx)
    If i = 0 Then
      nLast = eax
      If nLast = 0 Then
        nColor = GREY
      Else
        nColor = 0
      End If
    Else
      nColor = IIf(i > nLast, GREY, 0)
    End If

    Set itm = lvReg.ListItems.Add(, , "Std Level " & i)
    itm.ForeColor = nColor
    
    itm.SubItems(1) = cpu.HexPad(eax)
    itm.ListSubItems(1).ForeColor = nColor
    
    itm.SubItems(2) = cpu.HexPad(ebx)
    itm.ListSubItems(2).ForeColor = nColor
    
    itm.SubItems(3) = cpu.HexPad(ecx)
    itm.ListSubItems(3).ForeColor = nColor
    
    itm.SubItems(4) = cpu.HexPad(edx)
    itm.ListSubItems(4).ForeColor = nColor
  Next i
  
  For i = 0 To 8
    Call cpu.LevelGet(4 + i, eax, ebx, ecx, edx)
    If i = 0 Then
      nLast = eax
      If nLast = 0 Then
        nColor = GREY
      Else
        nColor = 0
        nLast = nLast And &HF
      End If
    Else
      nColor = IIf(i > nLast, GREY, 0)
    End If
    
    Set itm = lvReg.ListItems.Add(, , "Ext Level " & i)
    itm.ForeColor = nColor
    
    itm.SubItems(1) = cpu.HexPad(eax)
    itm.ListSubItems(1).ForeColor = nColor
    
    itm.SubItems(2) = cpu.HexPad(ebx)
    itm.ListSubItems(2).ForeColor = nColor
    
    itm.SubItems(3) = cpu.HexPad(ecx)
    itm.ListSubItems(3).ForeColor = nColor
    
    itm.SubItems(4) = cpu.HexPad(edx)
    itm.ListSubItems(4).ForeColor = nColor
  Next i
End Sub

'Decode Intel std cache descriptors
Private Sub StdCacheDescriptors(b As Byte, s1 As String, s2 As String, s3 As String, s4 As String, s5 As String, s6 As String)
  Select Case b
  Case &H1:   s1 = "Code TLB 4K":           s2 = "4 way":   s3 = "32"
  Case &H2:   s1 = "Code TLB 4M":           s2 = "fully":   s3 = "2"
  Case &H3:   s1 = "Data TLB 4K":           s2 = "4 way":   s3 = "64"
  Case &H4:   s1 = "Data TLB 4M":           s2 = "4 way":   s3 = "8"
  Case &H6:   s1 = "L1 code cache":         s2 = "4 way":   s4 = "8":       s5 = "32"
  Case &H8:   s1 = "L1 code cache":         s2 = "4 way":   s4 = "16":      s5 = "32"
  Case &HA:   s1 = "L1 data cache":         s2 = "2 way":   s4 = "8":       s5 = "32"
  Case &HC:   s1 = "L1 data cache":         s2 = "4 way":   s4 = "16":      s5 = "32"
  Case &H10:  s1 = "L1 data cache":         s2 = "4 way":   s4 = "16":      s5 = "32"
  Case &H15:  s1 = "L1 code cache":         s2 = "4 way":   s4 = "16":      s5 = "32"
  Case &H1A:  s1 = "L1 code/data":          s2 = "6 way":   s4 = "96":      s5 = "64"
  Case &H22:  s1 = "L3 code/data":          s2 = "4 way":   s4 = "512":     s5 = "64":  s6 = "dual"
  Case &H23:  s1 = "L3 code/data":          s2 = "8 way":   s4 = "1,024":   s5 = "64":  s6 = "dual"
  Case &H25:  s1 = "L3 code/data":          s2 = "8 way":   s4 = "2,048":   s5 = "64":  s6 = "dual"
  Case &H29:  s1 = "L3 code/data":          s2 = "8 way":   s4 = "4,096":   s5 = "64":  s6 = "dual"
  Case &H2C:  s1 = "L1 data cache":         s2 = "8 way":   s4 = "32":      s5 = "64"
  Case &H30:  s1 = "L1 code cache":         s2 = "8 way":   s4 = "32":      s5 = "64"
  Case &H39:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "128":     s5 = "64":  s6 = "yes"
  Case &H3B:  s1 = "L2 code/data":          s2 = "2 way":   s4 = "128":     s5 = "64":  s6 = "yes"
  Case &H3C:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "256":     s5 = "64":  s6 = "yes"
  Case &H40: 'no integrated L2 cache (P6 core) or L3 cache (P4 core)
  Case &H41:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "128":     s5 = "32"
  Case &H42:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "256":     s5 = "32"
  Case &H43:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "512":     s5 = "32"
  Case &H44:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "1,024":   s5 = "32"
  Case &H45:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "2,048":   s5 = "32"
  Case &H50:  s1 = "Code TLB 4K/4M/2M":     s2 = "fully":   s3 = "64"
  Case &H51:  s1 = "Code TLB 4K/4M/2M":     s2 = "fully":   s3 = "128"
  Case &H52:  s1 = "Code TLB 4K/4M/2M":     s2 = "fully":   s3 = "256"
  Case &H5B:  s1 = "Data TLB 4K/4M":        s2 = "fully":   s3 = "64"
  Case &H5C:  s1 = "Data TLB 4K/4M":        s2 = "fully":   s3 = "128"
  Case &H5D:  s1 = "Data TLB 4K/4M":        s2 = "fully":   s3 = "256"
  Case &H66:  s1 = "L1 data cache":         s2 = "4 way":   s4 = "8":       s5 = "64":  s6 = "yes"
  Case &H67:  s1 = "L1 data cache":         s2 = "4 way":   s4 = "16":      s5 = "64":  s6 = "yes"
  Case &H68:  s1 = "L1 data cache":         s2 = "4 way":   s4 = "32":      s5 = "64":  s6 = "yes"
  Case &H70:  s1 = "L1 trace cache":        s2 = "8 way":   s4 = "12 KµOps"
  Case &H71:  s1 = "L1 trace cache":        s2 = "8 way":   s4 = "16 KµOps"
  Case &H72:  s1 = "L1 trace cache":        s2 = "8 way":   s4 = "32 KµOps"
  Case &H77:  s1 = "L1 code cache":         s2 = "4 way":   s4 = "16":      s5 = "64":  s6 = "yes"
  Case &H79:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "128":     s5 = "64":  s6 = "dual"
  Case &H7A:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "256":     s5 = "64":  s6 = "dual"
  Case &H7B:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "512":     s5 = "64":  s6 = "dual"
  Case &H7C:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "1,024":   s5 = "64":  s6 = "dual"
  Case &H7E:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "256":     s5 = "128": s6 = "yes"
  Case &H81:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "128":     s5 = "32"
  Case &H82:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "256":     s5 = "32"
  Case &H83:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "512":     s5 = "32"
  Case &H84:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "1,024":   s5 = "32"
  Case &H85:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "2,048":   s5 = "32"
  Case &H86:  s1 = "L2 code/data":          s2 = "4 way":   s4 = "512":     s5 = "64"
  Case &H87:  s1 = "L2 code/data":          s2 = "8 way":   s4 = "1,024":   s5 = "64"
  Case &H88:  s1 = "L3 code/data":          s2 = "4 way":   s4 = "2,048":   s5 = "64"
  Case &H89:  s1 = "L3 code/data":          s2 = "4 way":   s4 = "4,096":   s5 = "64"
  Case &H8A:  s1 = "L3 code/data":          s2 = "4 way":   s4 = "8,192":   s5 = "64"
  Case &H8D:  s1 = "L3 code/data":          s2 = "12 way":  s4 = "3,096":   s5 = "128"
  Case &H90:  s1 = "Code TLB 4K..256M":     s2 = "fully":   s3 = "64"
  Case &H96:  s1 = "L1 data TLB 4K..256M":  s2 = "fully":   s3 = "32"
  Case &H9B:  s1 = "L2 data TLB 4K..256M":  s2 = "fully":   s3 = "96"
  Case &HB0:  s1 = "Code TLB 4K":           s2 = "4 way":   s3 = "128"
  Case &HB3:  s1 = "Data TLB 4K":           s2 = "4 way":   s3 = "128"
  Case Else
    Debug.Print "Unrecognised descriptor: " & b

  End Select
End Sub

'Format the application version number
Private Function FmtVersion() As String
  FmtVersion = App.Major & "." & Format$(App.Minor, "0#") & "." & Format$(App.Revision, "0###")
End Function

'Return the path from the passed path/filename.ext (Public for frmCapture)
Public Function GetPath(sFile As String) As String
  Dim i As Long
  
  i = InStrRev(sFile, "\")
  
  If i Then
    GetPath = Left$(sFile, i - 1)
  Else
    GetPath = ""
  End If
End Function

'Return the filename from the full path
Private Function GetFile(sFile As String) As String
  Dim i As Long
  
  i = InStrRev(sFile, "\")
  
  If i Then
    GetFile = Mid$(sFile, i + 1)
  Else
    GetFile = sFile
  End If
End Function

'Calculate the cpu speed, start the real time speed timer
Private Sub CalcSpeed()
  Dim i       As Long
  Dim fSecs   As Double
  Dim cCycles As Currency

  'It's not safe to calculate the full speed using a timer control because
  'SpeedStep cpu's clock down to a crawl if they're not active.
  Call cpu.BenchStart
    For i = 0 To 20000
      DoEvents
    Next i
  Call cpu.BenchStop(fSecs, cCycles)
  
  lblMHzFull.Caption = Format$(Round((cCycles / fSecs) / 1000000#, 0), "#,###") & " MHz"
  
'Start the real time speed calculation
  cpu.BenchStart
  tmrSpeed.Enabled = True
End Sub

'Autosize the passed ListView columns
Private Sub AutoSize(lv As ListView)
  Dim i As Long
  
  For i = 1 To lv.ColumnHeaders.Count
    Call SendMessage(lv.hWnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER)
  Next i
End Sub
