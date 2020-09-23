VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cIni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cTxtFile.cls - Simple ini file class
'
' Copyright free, use as and how you please.
' Paul_Caton@hotmail.com
'==============================================================================
Option Explicit

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As Any, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_sPath             As String
Private m_sKey              As String
Private m_sSection          As String
Private m_sDefault          As String
Private m_lLastReturnCode   As Long

Private Sub Class_Initialize()
  m_sPath = App.Path & "\" & App.Title & ".ini"
  m_sSection = "Setup"
End Sub

Public Property Get Default() As String
  Default = m_sDefault
End Property

Public Property Let Default(sDefault As String)
  m_sDefault = sDefault
End Property

Public Sub InvalidateCache()
  Call WritePrivateProfileString(vbNullString, vbNullString, vbNullString, m_sPath)
End Sub

Public Property Get Key() As String
  Key = m_sKey
End Property

Public Property Let Key(sKey As String)
  m_sKey = sKey
End Property

Public Property Get Path() As String
  Path = m_sPath
End Property

Public Property Let Path(sPath As String)
  m_sPath = sPath
End Property

Public Property Get Section() As String
  Section = m_sSection
End Property

Public Property Let Section(sSection As String)
  m_sSection = sSection
End Property

Public Property Get Value() As String
 Dim sBuf        As String
 Dim iSize       As Long
 Dim iRetCode    As Long

  sBuf = Space$(512)
  iSize = Len(sBuf)

  iRetCode = GetPrivateProfileString(m_sSection, m_sKey, m_sDefault, sBuf, iSize, m_sPath)

  If (iRetCode > 0) Then
        
    Value = Left$(sBuf, iRetCode)
   Else 'NOT (IRETCODE...
        
    Value = vbNullString
  End If
End Property

Public Property Let Value(ByVal sValue As String)
  sValue = Replace$(sValue, vbNullChar, vbNullString)
  m_lLastReturnCode = WritePrivateProfileString(m_sSection, m_sKey, sValue, m_sPath)
End Property
