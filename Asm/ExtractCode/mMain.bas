Attribute VB_Name = "mMain"
'==============================================================================
' Use this utility to extract the op-codes from an executable and store them
' in the clipboard as a string.
'
Option Explicit

Private Type OPENFILENAME
  lStructSize       As Long
  hWndOwner         As Long
  hInstance         As Long
  lpstrFilter       As String
  lpstrCustomFilter As String
  nMaxCustFilter    As Long
  nFilterIndex      As Long
  lpstrFile         As String
  nMaxFile          As Long
  lpstrFileTitle    As String
  nMaxFileTitle     As Long
  lpstrInitialDir   As String
  lpstrTitle        As String
  flags             As Long
  nFileOffset       As Integer
  nFileExtension    As Integer
  lpstrDefExt       As String
  lCustData         As Long
  lpfnHook          As Long
  lpTemplateName    As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Sub Main()
  Const PATH          As String = "Path"
  Dim OpenFileDialog  As OPENFILENAME
  Dim buf()           As Byte
  Dim nFile           As Integer
  Dim i               As Long
  Dim rv              As Long
  Dim nPos            As Long
  Dim nLen            As Long
  Dim sSrc            As String
  Dim sPatch          As String
  Dim sFile           As String
  
  With OpenFileDialog
    .lStructSize = Len(OpenFileDialog)
    .hInstance = App.hInstance
    .lpstrFilter = "Executables" + Chr$(0) + "*.exe"
    .lpstrFile = Space$(254)
    .nMaxFile = 255
    .lpstrFileTitle = Space$(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = GetSetting(App.ProductName, PATH, PATH, CurDir)
    .lpstrTitle = "Select executable..."
    .flags = 0
    
    rv = GetOpenFileName(OpenFileDialog)
   
    If rv = 0 Then Exit Sub
    
    Call SaveSetting(App.ProductName, PATH, PATH, GetPath(.lpstrFile))
    
    sFile = Left$(.lpstrFile, InStr(.lpstrFile, Chr$(0)) - 1)
  End With
  
  On Error GoTo Catch
  sSrc = ""
  nFile = FreeFile()
  Open sFile For Binary As #nFile
  Seek #nFile, &H1B1
  Get #nFile, , nLen
  Seek #nFile, 1024
  ReDim buf(0 To nLen)
  Get #nFile, , buf
  Close #nFile
  nFile = 0
  
  If nLen > 2048 Then
    Call MsgBox("Code too large", vbCritical)
    Exit Sub
  End If
  
  For i = 1 To nLen
    sSrc = sSrc & sHex(buf(i))
  Next i
  
  With Clipboard
    Call .Clear
    Call .SetText(sSrc)
  End With
  
  MsgBox "Hex-pair op-codes copied to the clipboard.", vbInformation
Catch:
End Sub

Private Function GetPath(ByVal sFile As String) As String
  Dim i As Long
  
  i = InStrRev(sFile, "\")
  
  If i Then
    GetPath = Left$(sFile, i - 1)
  Else
    GetPath = ""
  End If
End Function

Private Function sHex(a As Byte) As String
  sHex = Hex$(a)
  
  If Len(sHex) = 1 Then
    sHex = "0" & sHex
  End If
End Function
