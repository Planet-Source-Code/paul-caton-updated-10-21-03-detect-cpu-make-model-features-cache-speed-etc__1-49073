VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cCpuInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' cCpuInfo.cls - Class to host the results and interpretation of the CPUID
' op-code as implemented by the the cCPUID class
'
' Copyright free, use as and how you please.
' Paul_Caton@hotmail.com
'
' A brief intro to level support...
'   Pentium P2  Std Levels 0 to 2
'   Pentium P3  Std Levels 0 to 3
'   Pentium P4  Std Levels 0 to 2, Ext Levels 0, 2, 3, 4
'   AMD Athlon  Std Levels 0 to 1, Ext Levels 0 to 8
'
' A brief intro to levels...
'   Std 0 - Max std level supported and Vendor ID string
'   Std 1 - Model, family, type, features etc.
'   Std 2 - Cache descriptors (Intel)
'   Std 3 - Processor serial number (P3 only)
'
'   Ext 0 - Max extended level supported, Vendor ID string (AMD & Transmeta)
'   Ext 1 - Model, family, type, features etc. (AMD)
'   Ext 2 - Ext 4, processor name string
'   Ext 5 - L1 cache descriptors (AMD)
'   Ext 6 - L2 cache descriptors (AMD)
'   Ext 7 - Enhanced power management (AMD)
'   Ext 8 - Address size information (AMD)
'==============================================================================
Option Explicit

'Vendor enumeration
Public Enum eVendor
  ev_Unknown = 0
  ev_Intel
  ev_AMD
  ev_UMC
  ev_Cyrix
  ev_NexGen
  ev_Centaur
  ev_Rise
  ev_SIS
  ev_Transmeta
  ev_NatSemi
End Enum

'Feature enumeration
Public Enum eFeature
  'Level 1 ecx features
  ef_SSE3 = 0
  'ef_Reserved_1 = 1
  'ef_Reserved_2 = 2
  ef_MPN = 3
  ef_DFCPL = 4
  'ef_Reserved_5 = 5
  'ef_Reserved_6 = 6
  ef_EST = 7
  ef_TM2 = 8
  'ef_Reserved_9 = 9
  ef_CID = 10
  'ef_Reserved_C_11 = 11
  'ef_Reserved_C_12 = 12
  'ef_Reserved_C_13 = 13
  'ef_Reserved_C_14 = 14
  'ef_Reserved_C_15 = 15
  'ef_Reserved_C_16 = 16
  'ef_Reserved_C_17 = 17
  'ef_Reserved_C_18 = 18
  'ef_Reserved_C_19 = 19
  'ef_Reserved_C_20 = 20
  'ef_Reserved_C_21 = 21
  'ef_Reserved_C_22 = 22
  'ef_Reserved_C_23 = 23
  'ef_Reserved_C_24 = 24
  'ef_Reserved_C_25 = 25
  'ef_Reserved_C_26 = 26
  'ef_Reserved_C_27 = 27
  'ef_Reserved_C_28 = 28
  'ef_Reserved_C_29 = 29
  'ef_Reserved_C_30 = 30
  'ef_Reserved_C_31 = 31
  
  'Level 1 edx features
  ef_FPU = 32 + 0
  ef_VME = 32 + 1
  ef_DE = 32 + 2
  ef_PSE = 32 + 3
  ef_TSC = 32 + 4
  ef_MSR = 32 + 5
  ef_PAE = 32 + 6
  ef_MCE = 32 + 7
  ef_CX8 = 32 + 8
  ef_APIC = 32 + 9
  'ef_Reserved_D_10 = 32 + 10
  ef_SEP = 32 + 11
  ef_MTRR = 32 + 12
  ef_PGE = 32 + 13
  ef_MCA = 32 + 14
  ef_CMOV = 32 + 15
  ef_PAT = 32 + 16
  ef_PSE36 = 32 + 17
  ef_PSN = 32 + 18
  ef_CLFSH = 32 + 19
  'ef_Reserved_D_20 = 32 + 20
  ef_DS = 32 + 21
  ef_ACPI = 32 + 22
  ef_MMX = 32 + 23
  ef_FXSR = 32 + 24
  ef_SSE = 32 + 25
  ef_SSE2 = 32 + 26
  ef_SS = 32 + 27
  ef_HTT = 32 + 28
  ef_TM1 = 32 + 29
  ef_IA64 = 32 + 30
  ef_PBE = 32 + 31
  
  'Extended Level 1 edx features
  ef_xFPU = 64 + 0
  ef_xVME = 64 + 1
  ef_xDE = 64 + 2
  ef_xPSE = 64 + 3
  ef_xTSC = 64 + 4
  ef_xMSR = 64 + 5
  ef_xPAE = 64 + 6
  ef_xMCE = 64 + 7
  ef_xCX8 = 64 + 8
  ef_xAPIC = 64 + 9
  'ef_xReserved_10 = 64 + 10
  ef_xSEP = 64 + 11
  ef_xMTRR = 64 + 12
  ef_xPGE = 64 + 13
  ef_xMCA = 64 + 14
  ef_xCMOV = 64 + 15
  ef_xPAT = 64 + 16
  ef_xFCMOV = 64 + 16
  ef_xPSE36 = 64 + 17
  'ef_xReserved_18 = 18
  ef_xMP = 64 + 19
  ef_xNX = 64 + 20
  'ef_xReserved_21 = 64 + 21
  ef_xMMX_PLUS = 64 + 22
  ef_xMMX = 64 + 23
  ef_xMMX_PLUS_CYRIX = 64 + 24
  ef_xFXSR = 64 + 24
  'ef_xReserved_25 = 25
  'ef_xReserved_26 = 26
  'ef_xReserved_27 = 27
  'ef_xReserved_28 = 28
  ef_xLM = 64 + 29
  ef_x3DNOW_X = 64 + 30
  ef_x3DNOW = 64 + 31
End Enum

'Processor type enumeration
Public Enum eProcessorType
  ept_Primary_Processor = 0
  ept_Overdrive_Processor
  ept_Secondary_Processor
  ept_Reserved
End Enum

'Serial number status enumeration
Public Enum eSnStatus
  esn_Unavailable = 0
  esn_Disabled
  esn_Enabled
End Enum

'Register index enumeration
Public Enum eLevels
  el_Std0 = 0
  el_Std1
  el_Std2
  el_Std3
  el_Xtd0
  el_Xtd1
  el_Xtd2
  el_Xtd3
  el_Xtd4
  el_Xtd5
  el_Xtd6
  el_Xtd7
  el_Xtd8
End Enum

'Extended: Cache fields enumeration
Public Enum eExCacheFields
  ecf_L1_4MbDataAssoc
  ecf_L1_4MbDataEntries
  ecf_L1_4MbCodeAssoc
  ecf_L1_4MbCodeEntries
  
  ecf_L1_4KbDataAssoc
  ecf_L1_4KbDataEntries
  ecf_L1_4KbCodeAssoc
  ecf_L1_4KbCodeEntries
  
  ecf_L1_DataSizeKb
  ecf_L1_DataAssoc
  ecf_L1_DataLinesPerTag
  ecf_L1_DataLineSizeBytes
  
  ecf_L1_CodeSizeKb
  ecf_L1_CodeAssoc
  ecf_L1_CodeLinesPerTag
  ecf_L1_CodeLineSizeBytes
  
  ecf_L2_4MbDataAssoc
  ecf_L2_4MbDataEntries
  ecf_L2_4MbCodeAssoc
  ecf_L2_4MbCodeEntries
  
  ecf_L2_4KbDataAssoc
  ecf_L2_4KbDataEntries
  ecf_L2_4KbCodeAssoc
  ecf_L2_4KbCodeEntries
  
  ecf_L2_CacheSizeKb
  ecf_L2_CacheAssoc
  ecf_L2_CacheLinesPerTag
  ecf_L2_CacheLineSizeBytes
End Enum

'Extended: power management enumeration
Public Enum eExPowerManagment
  epm_TemperatureSensor = 0
  epm_FrequencyID = 1
  epm_VoltageID = 2
  epm_ThermalTrip = 3
  epm_ThermalMonitoring = 4
  epm_SoftwareThermalControl = 5
End Enum

'Api declarations
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long

'Vendor ID constants
Private Const VEND_INTEL      As String = "GenuineIntel"
Private Const VEND_UMC        As String = "UMC UMC UMC "
Private Const VEND_AMD        As String = "AuthenticAMD"
Private Const VEND_CYRIX      As String = "CyrixInstead"
Private Const VEND_NEXGEN     As String = "NexGenDriven"
Private Const VEND_CENTAUR    As String = "CentaurHauls"
Private Const VEND_RISE       As String = "RiseRiseRise"
Private Const VEND_SIS        As String = "SiS SiS SiS "
Private Const VEND_TRANSMETA  As String = "GenuineTMx86"
Private Const VEND_NAT_SEMI   As String = "Geode by NSC"

'Extended processor family enumeration
Private Enum eExFamily
  Intel_P4_Amd_K8 = 0
  Intel_Itanium2
End Enum

'Registers type
Private Type tRegisters
  eax As Long                                               'Cpu eax register
  ebx As Long                                               'Cpu ebx register
  ecx As Long                                               'Cpu ecx register
  edx As Long                                               'Cpu edx register
End Type

'Private member variables
Private m_nLevelsStd  As Long                               'Max std levels
Private m_nLevelsExt  As Long                               'Max ext levels
Private m_nVendor     As eVendor                            'Vendor

Private cCpuStart     As Currency                           'BenchStart cpu clock cycles
Private cCpuStop      As Currency                           'BenchStop cpu clock cycles
Private cQpcFreq      As Currency                           'QueryPerformance frequency
Private cQpcStart     As Currency                           'QueryPerformance start
Private cQpcStop      As Currency                           'QueryPerformance stop

Private Std(0 To 3)   As tRegisters                         'Std registers
Private Ext(0 To 8)   As tRegisters                         'Ext registers
Private cpu           As cCPUID                             'cCPUID instance

Private Sub Class_Initialize()
  On Error Resume Next
  Set cpu = New cCPUID
  Call Refresh(True)
  Call QueryPerformanceFrequency(cQpcFreq)                  'Get the QueryPerformance frequency
End Sub

Private Sub Class_Terminate()
  Set cpu = Nothing
End Sub

'APIC ID
Public Property Get ApicID() As Long
  If m_nLevelsStd >= 1 Then
    ApicID = cpu.ShiftRight(Std(1).ebx, 24) And &HFF
  End If
End Property

'Benchmark start
Public Sub BenchStart()
  Call QueryPerformanceCounter(cQpcStart)
  Call cpu.CpuClk(cCpuStart)
End Sub

'Benchmark stop
Public Sub BenchStop(fSeconds As Double, cCpuCycles As Currency)
  Call QueryPerformanceCounter(cQpcStop)
  Call cpu.CpuClk(cCpuStop)
  
  fSeconds = CDbl((cQpcStop - cQpcStart) / cQpcFreq)
  cCpuCycles = (cCpuStop - cCpuStart) * 10000
End Sub

'Return the std cache descriptors as a byte array
Public Property Get CacheDescriptors() As Byte()
  Dim nCount    As Byte
  Dim bArray()  As Byte
  Dim i         As Long
  Dim j         As Long
  Dim s         As String
  Dim t         As String
  
  If m_nLevelsStd >= 2 Then
    With Std(2)
      nCount = .eax And &HFF                                'The low byte indicates how many times we need to call CPUID level 2
      For i = 1 To nCount
        If cpu.ShiftRight(.eax, 31) = 0 Then                'Most significant bit must be 0
          s = s & HexPad(.eax And &HFFFFFF00)               'Mask out the count field in the conversion
        End If
        
        If cpu.ShiftRight(.ebx, 31) = 0 Then                'Most significant bit must be 0
          s = s & HexPad(.ebx)
        End If
  
        If cpu.ShiftRight(.ecx, 31) = 0 Then                'Most significant bit must be 0
          s = s & HexPad(.ecx)
        End If
          
        If cpu.ShiftRight(.edx, 31) = 0 Then                'Most significant bit must be 0
          s = s & HexPad(.edx)
        End If
        
        If nCount > 1 Then                                  'If the count is greater than 1
          Call cpu.CpuID(2, .eax, .ebx, .ecx, .edx)         'Read level 2 again
        End If
      Next i
    End With
    
    'Convert the string to a byte array
    j = 1
    For i = 1 To Len(s) Step 2
      t = Mid$(s, i, 2)
      If t <> "00" Then                                       'If not a null descriptor
        ReDim Preserve bArray(1 To j)
        bArray(j) = Val("&H" & t)
        j = j + 1
      End If
    Next i
  Else
    ReDim bArray(1 To 1)
    bArray(1) = 0
  End If
  
  CacheDescriptors = bArray
End Property

'Return the cpu brand
Public Function CpuBrand(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuBrand = Std(1).ebx
    End If
  Else
    'Only AMD Opteron at this point in time
    If m_nLevelsExt >= 1 Then
      CpuBrand = Ext(1).ebx And &HFFF
    End If
  End If
End Function

'Return the cpu brand as a string
Public Function CpuBrandStr(Optional Standard As Boolean = True) As String
  Dim nCode   As Long
  Dim nSpeed  As Long
  Dim nBrand  As Long
  
  If Standard Then
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  Else
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  End If
  
  If Standard Then
    Select Case CpuBrand(True)
    Case &H1:   CpuBrandStr = "0.18 µm Intel Celeron"
    Case &H2:   CpuBrandStr = "0.18 µm Intel Pentium III"
    Case &H3:   CpuBrandStr = "0.18 µm Intel Pentium III Xeon OR 0.13 µm Intel Celeron"
    Case &H4:   CpuBrandStr = "0.13 µm Intel Pentium III"
    Case &H6:   CpuBrandStr = "0.13 µm Intel Pentium III mobile"
    Case &H7:   CpuBrandStr = "0.13 µm Intel Celeron mobile"
    Case &H8:   CpuBrandStr = "0.18 µm Intel Pentium 4 OR 0.13 µm Intel Celeron 4 mobile (0F24h)"
    Case &H9:   CpuBrandStr = "0.13 µm Intel Pentium 4"
    Case &HA:   CpuBrandStr = "0.18 µm Intel Celeron 4"
    Case &HB:   CpuBrandStr = "0.18 µm Intel Pentium 4 Xeon MP OR 0.13 µm Intel Pentium 4 Xeon"
    Case &HC:   CpuBrandStr = "0.13 µm Intel Pentium 4 Xeon MP"
    Case &HE:   CpuBrandStr = "0.13 µm Intel Pentium 4 mobile"
    Case &HF:   CpuBrandStr = "0.13 µm Intel Celeron 4 mobile (0F27h)"
    Case &H16:  CpuBrandStr = "0.13 µm Intel Pentium M"
    Case Else
      If m_nVendor = ev_AMD Then
        nBrand = CpuBrand(True)
        nCode = cpu.ShiftRight(nCode, 5)
        nSpeed = nBrand And &H1F
        nSpeed = 32 + (2 * nSpeed)
        
        Select Case nCode
        Case &H3:   CpuBrandStr = "Opteron UP 1" & nSpeed
        Case &H4:   CpuBrandStr = "Opteron DP 2" & nSpeed
        Case &H5:   CpuBrandStr = "Opteron MP 8" & nSpeed
        End Select
      End If
    End Select
  Else
    If m_nVendor = ev_AMD Then
      nBrand = CpuBrand(False)
      If nBrand <> 0 Then
        nCode = cpu.ShiftRight(nBrand, 8)
        nSpeed = nBrand And &HFF
        nSpeed = 32 + (2 * nSpeed)
        Select Case nCode
        Case &H3:   CpuBrandStr = "Opteron UP 1" & nSpeed
        Case &H4:   CpuBrandStr = "Opteron DP 2" & nSpeed
        Case &H5:   CpuBrandStr = "Opteron MP 8" & nSpeed
        End Select
      End If
    End If
  End If
End Function

'Return the number of logical cpu count. Hyper-Threading cpu's will return 2
Public Property Get CpuLogicalCount()
  If m_nLevelsStd >= 1 Then
    CpuLogicalCount = cpu.ShiftRight(Std(1).ebx, 16) And &HFF
  End If
End Property

'Return the cpu family
Public Function CpuFamily(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuFamily = cpu.ShiftRight(Std(1).eax, 8) And &HF
    End If
  Else
    If m_nLevelsExt >= 1 Then
      CpuFamily = cpu.ShiftRight(Ext(1).eax, 8) And &HF
    End If
  End If
End Function

'Return the cpu family ex
Public Function CpuFamilyEx(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuFamilyEx = cpu.ShiftRight(Std(1).eax, 20) And &HFF
    End If
  Else
    If m_nLevelsExt >= 1 Then
      CpuFamilyEx = cpu.ShiftRight(Ext(1).eax, 20) And &HFF
    End If
  End If
End Function

'Return the cpu family as a string
Public Function CpuFamilyStr(Optional Standard As Boolean = True) As String
  Dim nFamily   As Long
  Dim nFamilyEx As Long
  
  If Standard Then
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  Else
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  End If
  
  nFamily = CpuFamily(Standard)
  nFamilyEx = CpuFamilyEx(Standard)
  
  If Standard Then
    Select Case nFamily
    Case &H4
      Select Case m_nVendor
      Case eVendor.ev_Intel:          CpuFamilyStr = "80486"
      Case eVendor.ev_AMD:            CpuFamilyStr = "5x86"
      Case eVendor.ev_Cyrix:          CpuFamilyStr = "5x86"
      End Select
    
    Case &H5
      Select Case m_nVendor
      Case eVendor.ev_Intel:          CpuFamilyStr = "P5, P54C, P55C, P24T"
      Case eVendor.ev_AMD:            CpuFamilyStr = "K5, K6"
      Case eVendor.ev_Cyrix:          CpuFamilyStr = "M1"
      Case eVendor.ev_NexGen:         CpuFamilyStr = "Nx586"
      Case eVendor.ev_Centaur:        CpuFamilyStr = "C6, C2, C3"
      Case eVendor.ev_Rise:           CpuFamilyStr = "mP6"
      Case eVendor.ev_SIS:            CpuFamilyStr = "55x"
      Case eVendor.ev_Transmeta:      CpuFamilyStr = "TM3x00, TM5x00"
      End Select
    
    Case &H6
      Select Case m_nVendor
      Case eVendor.ev_Intel:          CpuFamilyStr = "P6, P2, P3, PM"
      Case eVendor.ev_AMD:            CpuFamilyStr = "K7"
      Case eVendor.ev_Cyrix:          CpuFamilyStr = "M2"
      Case eVendor.ev_Centaur:        CpuFamilyStr = "C3"
      End Select
    
    Case &H7
      Select Case m_nVendor
      Case eVendor.ev_Intel:          CpuFamilyStr = "Itanium"
      End Select
    
    Case &HF
      Select Case nFamilyEx
      Case eExFamily.Intel_P4_Amd_K8:
        Select Case m_nVendor
        Case eVendor.ev_Intel:        CpuFamilyStr = "P4"
        Case eVendor.ev_AMD:          CpuFamilyStr = "K8"
        End Select
      
      Case eExFamily.Intel_Itanium2:  CpuFamilyStr = "Itanium 2"
      End Select
    End Select
  Else
    'Intel doesn't use the extended family, model
    Select Case nFamily
    Case &H5
      Select Case m_nVendor
      Case eVendor.ev_AMD:            CpuFamilyStr = "K5"
      Case eVendor.ev_Centaur:        CpuFamilyStr = "C2, C3"
      Case eVendor.ev_Transmeta:      CpuFamilyStr = "TM3x00, TM5x00"
      End Select
    
    Case &H6
      Select Case m_nVendor
      Case eVendor.ev_AMD:            CpuFamilyStr = "K6"
      Case eVendor.ev_Centaur:        CpuFamilyStr = "C3"
      End Select
    
    Case &H7
      Select Case m_nVendor
      Case eVendor.ev_AMD:            CpuFamilyStr = "K7"
      End Select
    
    Case &HF
      Select Case m_nVendor
      Case eVendor.ev_AMD:
        Select Case nFamilyEx
        Case 0:                       CpuFamilyStr = "K8"
        End Select
      End Select
    End Select
  End If
End Function

'Return the cpu model
Public Function CpuModel(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuModel = cpu.ShiftRight(Std(1).eax, 4) And &HF
    End If
  Else
    If m_nLevelsExt >= 1 Then
      CpuModel = cpu.ShiftRight(Ext(1).eax, 4) And &HF
    End If
  End If
End Function

'Return the cpu model ex
Public Function CpuModelEx(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuModelEx = cpu.ShiftRight(Std(1).eax, 16) And &HF
    End If
  Else
    If m_nLevelsExt >= 1 Then
      CpuModelEx = cpu.ShiftRight(Ext(1).eax, 16) And &HF
    End If
  End If
End Function

'Return the cpu model as a string
Public Function CpuModelStr(Optional Standard As Boolean = True) As String
  Dim nCpuModel     As Long
  Dim nCpuFamily    As Long
  Dim nCpuFamilyEx  As Long
  
  If Standard Then
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  Else
    If m_nLevelsStd < 1 Then
      Exit Function
    End If
  End If
  
  nCpuModel = CpuModel(Standard)
  nCpuFamily = CpuFamily(Standard)
  nCpuFamilyEx = CpuFamilyEx(Standard)
  
  If Standard Then
    Select Case m_nVendor
    Case eVendor.ev_Intel
      Select Case nCpuFamily
      Case &H4
        Select Case nCpuModel
        Case &H0:   CpuModelStr = "i80486DX-25/33"
        Case &H1:   CpuModelStr = "i80486DX-50"
        Case &H2:   CpuModelStr = "i80486SX"
        Case &H3:   CpuModelStr = "i80486DX2"
        Case &H4:   CpuModelStr = "i80486SL"
        Case &H5:   CpuModelStr = "i80486SX2"
        Case &H7:   CpuModelStr = "i80486DX2WB"
        Case &H8:   CpuModelStr = "i80486DX4"
        Case &H9:   CpuModelStr = "i80486DX4WB"
        End Select
      
      Case &H5
        Select Case nCpuModel
        Case &H0:   CpuModelStr = "P5 A-step"
        Case &H1:   CpuModelStr = "P5"
        Case &H2:   CpuModelStr = "P54C"
        Case &H3:   CpuModelStr = "P24T Overdrive"
        Case &H4:   CpuModelStr = "P55C"
        Case &H7:   CpuModelStr = "P55C"
        Case &H8:   CpuModelStr = "P55C (0.25µm)"
        End Select
      
      Case &H6
        Select Case nCpuModel
        Case &H0:   CpuModelStr = "P6 A-step"
        Case &H1:   CpuModelStr = "P6"
        Case &H3:   CpuModelStr = "P2 (0.28 µm)"
        Case &H5:   CpuModelStr = "P2 (0.25 µm)"
        Case &H6:   CpuModelStr = "P2 with on-die L2 cache"
        Case &H7:   CpuModelStr = "P3 (0.25 µm)"
        Case &H8:   CpuModelStr = "P3 (0.18 µm) with 256 KB on-die L2 cache"
        Case &H9:   CpuModelStr = "PM (0.13 µm) with 1 MB on-die L2 cache"
        Case &HA:   CpuModelStr = "P3 (0.18 µm) with 1 or 2 MB on-die L2 cache"
        Case &HB:   CpuModelStr = "P3 (0.13 µm) with 256 or 512 KB on-die L2 cache"
        End Select
      
      Case &H7
      
      Case &HF
        Select Case nCpuFamilyEx
        Case eExFamily.Intel_P4_Amd_K8:
          Select Case nCpuModel
          Case 0:   CpuModelStr = "P4 (0.18 µm)"
          Case 1:   CpuModelStr = "P4 (0.18 µm)"
          Case 2:   CpuModelStr = "P4 (0.13 µm)"
          Case 3:   CpuModelStr = "P4 (0.09 µm)"
          End Select
        
        Case eExFamily.Intel_Itanium2:
        End Select
      End Select
  
    Case eVendor.ev_AMD
      Select Case nCpuFamily
      Case &H4
        Select Case nCpuModel
        Case &H3:   CpuModelStr = "80486DX2"
        Case &H7:   CpuModelStr = "80486DX2WB"
        Case &H8:   CpuModelStr = "80486DX4"
        Case &H9:   CpuModelStr = "80486DX4WB"
        Case &HE:   CpuModelStr = "5x86"
        Case &HF:   CpuModelStr = "5x86WB"
        End Select
      
      Case &H5
        Select Case nCpuModel
        Case &H0:   CpuModelStr = "SSA5 (PR75, PR90, PR100)"
        Case &H1:   CpuModelStr = "5k86 (PR120, PR133)"
        Case &H2:   CpuModelStr = "5k86 (PR166)"
        Case &H3:   CpuModelStr = "5k86 (PR200)"
        Case &H6:   CpuModelStr = "K6 (0.30 µm)"
        Case &H7:   CpuModelStr = "K6 (0.25 µm)"
        Case &H8:   CpuModelStr = "K6-2"
        Case &H9:   CpuModelStr = "K6-III"
        Case &HD:   CpuModelStr = "K6-2+ or K6-III+ (0.18 µm)"
        End Select
      
      Case &H6
        Select Case nCpuModel
        Case &H1:   CpuModelStr = "Athlon (0.25 µm)"
        Case &H2:   CpuModelStr = "Athlon (0.18 µm)"
        Case &H3:   CpuModelStr = "Duron (SF core)"
        Case &H4:   CpuModelStr = "Athlon (TB core)"
        Case &H6:   CpuModelStr = "Athlon (PM core)"
        Case &H7:   CpuModelStr = "Duron (MG core)"
        Case &H8:   CpuModelStr = "Athlon (TH core)"
        Case &HA:   CpuModelStr = "Athlon (BT core)"
        End Select
      
      Case &H7
      
      Case &HF
        Select Case nCpuFamilyEx
        Case eExFamily.Intel_P4_Amd_K8:
          Select Case nCpuModel
          Case 5:   CpuModelStr = "Opteron DP (0.13 µm)"
          End Select
        Case eExFamily.Intel_Itanium2:
        End Select
      End Select
  
    Case eVendor.ev_Centaur
      Select Case nCpuModel
      Case &H4:     CpuModelStr = "C6"
      Case &H8:     CpuModelStr = "C2"
      Case &H9:     CpuModelStr = "C3"
      End Select
      
    Case eVendor.ev_Cyrix
      Select Case nCpuFamily
      Case &H4:
        Select Case nCpuModel
        Case &H4:   CpuModelStr = "GX, GXm"
        Case &H9:   CpuModelStr = "5x86"
        End Select
      
      Case &H5:
        Select Case nCpuModel
        Case &H2:   CpuModelStr = "6x86"
        Case &H5:   CpuModelStr = "Cyrix M2 core"
        Case &H6:   CpuModelStr = "WinChip C5A core"
        Case &H7
          If CpuStepping(Standard) < 8 Then
            CpuModelStr = "WinChip C5B core"
          Else
            CpuModelStr = "WinChip C5C core"
          End If
        Case &H8:   CpuModelStr = "WinChip C5N core"
        Case &H9:   CpuModelStr = "WinChip C5XL core"
        End Select
      
      Case &H6:
        Select Case nCpuModel
        Case &H0:   CpuModelStr = "6x86MX"
        End Select
      End Select
      
    Case eVendor.ev_NexGen
      Select Case nCpuModel
      Case &H0:      CpuModelStr = "Nx586 or Nx586FPU"
      End Select
      
    Case eVendor.ev_Rise
      Select Case nCpuModel
      Case &H0:      CpuModelStr = "mP6 (0.25 µm)"
      Case &H2:      CpuModelStr = "mP6 (0.18 µm)"
      End Select
      
    Case eVendor.ev_SIS
      Select Case nCpuModel
      Case &H0:      CpuModelStr = "55x"
      End Select
      
    Case eVendor.ev_Transmeta
      Select Case nCpuModel
      Case &H4:      CpuModelStr = "Crusoe TM3x00 and TM5x00"
      End Select
      
    Case eVendor.ev_UMC
      Select Case nCpuModel
      Case &H1:      CpuModelStr = "U5D"
      Case &H2:      CpuModelStr = "U5S"
      End Select
    End Select
  Else
    'Intel doesn't use the extended family, model
    Select Case m_nVendor
    Case eVendor.ev_AMD
      Select Case nCpuFamily
      Case &H5
        Select Case nCpuModel
        Case &H1:   CpuModelStr = "5k86 (PR120 or PR133)"
        Case &H2:   CpuModelStr = "5k86 (PR166)"
        Case &H3:   CpuModelStr = "5k86 (PR200)"
        End Select
      
      Case &H6
        Select Case nCpuModel
        Case &H6:   CpuModelStr = "K6 (0.30 µm)"
        Case &H7:   CpuModelStr = "K6 (0.25 µm)"
        Case &H8:   CpuModelStr = "K6-2"
        Case &H9:   CpuModelStr = "K6-III"
        Case &HD:   CpuModelStr = "K6-2+ or K6-III+ (0.18 µm)"
        End Select
      
      Case &H7
        Select Case nCpuModel
        Case &H1:   CpuModelStr = "Athlon (0.25 µm)"
        Case &H2:   CpuModelStr = "Athlon (0.18 µm)"
        Case &H3:   CpuModelStr = "Duron (SF core)"
        Case &H4:   CpuModelStr = "Athlon (TB core)"
        Case &H6:   CpuModelStr = "Athlon (PM core)"
        Case &H7:   CpuModelStr = "Duron (MG core)"
        Case &H8:   CpuModelStr = "Athlon (TH core)"
        Case &HA:   CpuModelStr = "Athlon (BT core)"
        End Select
      
      Case &HF
        Select Case nCpuFamilyEx
        Case 0:
          Select Case nCpuModel
          Case &H5: CpuModelStr = "Opteron DP (0.13 µm)"
          End Select
        End Select
      End Select
  
    Case eVendor.ev_Centaur
      Select Case nCpuModel
      Case &H8:     CpuModelStr = "C2"
      Case &H9:     CpuModelStr = "C3"
      End Select
      
    Case eVendor.ev_Cyrix
      Select Case nCpuModel
      Case &H6:     CpuModelStr = "WinChip C5A core"
      Case &H7
        If CpuStepping(Standard) < 8 Then
                    CpuModelStr = "WinChip C5B core"
        Else
                    CpuModelStr = "WinChip C5C core"
        End If
      
      Case &H8:     CpuModelStr = "WinChip C5N core"
      Case &H9:     CpuModelStr = "WinChip C5XL core"
      End Select
      
    Case eVendor.ev_Transmeta
      Select Case nCpuModel
      Case &H4:      CpuModelStr = "Crusoe TM3x00 and TM5x00"
      End Select
    End Select
  End If
  
  If nCpuModel = &HF Then
    CpuModelStr = CpuModelStr & " - Extended model #" & Hex$(CpuModelEx(Standard))
  End If
End Function

'Return the cpu stepping
Public Function CpuStepping(Optional Standard As Boolean = True) As Long
  If Standard Then
    If m_nLevelsStd >= 1 Then
      CpuStepping = Std(1).eax And &HF
    End If
  Else
    If m_nLevelsExt >= 1 Then
      CpuStepping = Ext(1).eax And &HF
    End If
  End If
End Function

'Return the cpu type
Public Property Get CpuType() As eProcessorType
  If m_nLevelsStd >= 1 Then
    CpuType = cpu.ShiftRight(Std(1).eax, 12) And &H3
  End If
End Property

'Return the cpu type as a string
Public Property Get CpuTypeStr() As String
  If m_nLevelsStd >= 1 Then
    Select Case cpu.ShiftRight(Std(1).eax, 12) And &H3
    Case eProcessorType.ept_Primary_Processor:    CpuTypeStr = "Primary processor"
    Case eProcessorType.ept_Overdrive_Processor:  CpuTypeStr = "Overdrive processor"
    Case eProcessorType.ept_Secondary_Processor:  CpuTypeStr = "Secondary processor"
    End Select
  End If
End Property

'Return whether extended levels are supported
Public Property Get Extended() As Boolean
  Extended = m_nLevelsExt <> 0
End Property

'Return whether a feature is supported
Public Function Feature(nType As eFeature) As Boolean
  Select Case nType
  Case 0 To 31
    If m_nLevelsStd >= 1 Then
      Feature = cpu.ShiftRight(Std(1).ecx, nType) And 1
    End If
    
  Case 32 To 63
    If m_nLevelsStd >= 1 Then
      Feature = cpu.ShiftRight(Std(1).edx, nType - 32) And 1
    End If
    
  Case 64 To 95
    If m_nLevelsExt >= 1 Then
        Feature = cpu.ShiftRight(Ext(1).edx, nType - 64) And 1
    End If
    
  End Select
End Function

'Utility function, return the passed value as a hex string with prepended leading zeros to the passed width
Public Function HexPad(nValue As Long, Optional nWidth As Long = 8) As String
  Dim nLen As Long
  
  HexPad = Hex$(nValue)
  nLen = Len(HexPad)
  
  If nLen < nWidth Then
    HexPad = String$(nWidth - Len(HexPad), "0") & HexPad
  End If
End Function

'Return the maximum supported extended level
Public Property Get LevelsExt() As Long
  LevelsExt = m_nLevelsExt
End Property

'Return the maximum supported standard level
Public Property Get LevelsStd() As Long
  LevelsStd = m_nLevelsStd
End Property

'Return the name string
Public Property Get NameStr() As String
  Dim i As Long
  
  If m_nLevelsExt >= 4 Then
    For i = 2 To 4
      NameStr = NameStr & _
                DecodeValue(Ext(i).eax) & _
                DecodeValue(Ext(i).ebx) & _
                DecodeValue(Ext(i).ecx) & _
                DecodeValue(Ext(i).edx)
    Next i
  End If
  
  NameStr = Trim$(NameStr)
End Property

'Get the register values at the passed level so user can get at the raw data
Public Sub LevelGet(nReg As eLevels, eax As Long, ebx As Long, ecx As Long, edx As Long)
  Select Case nReg
  Case eLevels.el_Std0 To eLevels.el_Std3
    With Std(nReg)
      eax = .eax
      ebx = .ebx
      ecx = .ecx
      edx = .edx
    End With
    
  Case eLevels.el_Xtd0 To eLevels.el_Xtd8
    With Ext(nReg - 4)
      eax = .eax
      ebx = .ebx
      ecx = .ecx
      edx = .edx
    End With
    
  End Select
End Sub

'Set the register values at the passed level
Public Sub LevelSet(nReg As eLevels, eax As Long, ebx As Long, ecx As Long, edx As Long)
  Select Case nReg
  Case eLevels.el_Std0 To eLevels.el_Std3
    With Std(nReg)
      .eax = eax
      .ebx = ebx
      .ecx = ecx
      .edx = edx
    End With
    
  Case eLevels.el_Xtd0 To eLevels.el_Xtd8
    With Ext(nReg - 4)
      .eax = eax
      .ebx = ebx
      .ecx = ecx
      .edx = edx
    End With
    
  End Select
End Sub

'Return the manufacturer as a string
Public Property Get ManufacturerStr() As String
  If m_nLevelsStd > 0 Then
    Select Case m_nVendor
    Case eVendor.ev_Intel:      ManufacturerStr = "Intel"
    Case eVendor.ev_AMD:        ManufacturerStr = "AMD"
    Case eVendor.ev_UMC:        ManufacturerStr = "UMC"
    Case eVendor.ev_Cyrix:      ManufacturerStr = "Cyrix"
    Case eVendor.ev_NexGen:     ManufacturerStr = "NexGen"
    Case eVendor.ev_Centaur:    ManufacturerStr = "Centaur"
    Case eVendor.ev_Rise:       ManufacturerStr = "Rise"
    Case eVendor.ev_SIS:        ManufacturerStr = "SIS"
    Case eVendor.ev_Transmeta:  ManufacturerStr = "Transmeta"
    Case eVendor.ev_NatSemi:    ManufacturerStr = "National Semiconductor"
    Case Else:                  ManufacturerStr = "Unknown"
    End Select
  End If
End Property

Public Sub Refresh(bReReadCPU As Boolean)
  Dim i         As Long
  Dim sVendorID As String
  
  If bReReadCPU Then
    'Read the std levels into the registers, for simplicity... read them all
    For i = 0 To 3
      With Std(i)
        Call cpu.CpuID(i, .eax, .ebx, .ecx, .edx)
      End With
    Next i
    
    'Read the ext levels into the registers, for simplicity... read them all
    For i = 0 To 8
      With Ext(i)
        Call cpu.CpuID(&H80000000 + i, .eax, .ebx, .ecx, .edx)
      End With
    Next i
  Else
    'Don't read from cpu, use the register values as they are...
    'Perhaps the registers have been set using LevelSet
  End If
  
  m_nLevelsStd = Std(0).eax And &HFF                        'Store the number of std levels used
  
  If m_nLevelsStd > 0 Then
    m_nLevelsExt = Ext(0).eax And &HFF                      'Store the number of ext levels used
    
    'Decode the Vendor ID string
    sVendorID = VendorIdStr(True)
    
    'Deduce the vendor from the Vendor ID string
    Select Case 0
    Case StrComp(sVendorID, VEND_INTEL):      m_nVendor = eVendor.ev_Intel
    Case StrComp(sVendorID, VEND_AMD):        m_nVendor = eVendor.ev_AMD
    Case StrComp(sVendorID, VEND_UMC):        m_nVendor = eVendor.ev_UMC
    Case StrComp(sVendorID, VEND_CYRIX):      m_nVendor = eVendor.ev_Cyrix
    Case StrComp(sVendorID, VEND_NEXGEN):     m_nVendor = eVendor.ev_NexGen
    Case StrComp(sVendorID, VEND_CENTAUR):    m_nVendor = eVendor.ev_Centaur
    Case StrComp(sVendorID, VEND_RISE):       m_nVendor = eVendor.ev_Rise
    Case StrComp(sVendorID, VEND_SIS):        m_nVendor = eVendor.ev_SIS
    Case StrComp(sVendorID, VEND_TRANSMETA):  m_nVendor = eVendor.ev_Transmeta
    Case StrComp(sVendorID, VEND_NAT_SEMI):   m_nVendor = eVendor.ev_NatSemi
    Case Else:                                m_nVendor = eVendor.ev_Unknown
    End Select
  End If
End Sub

'Return the serial number status
Public Property Get SnStatus() As eSnStatus
  If m_nLevelsStd < 3 Then
    SnStatus = esn_Unavailable
  Else
    If Feature(ef_PSN) Then
      SnStatus = esn_Enabled
    Else
      SnStatus = esn_Disabled
    End If
  End If
End Property

'Return the 96 bit serial number formatted in standard fashion
Public Property Get SnStr() As String
  If SnStatus = esn_Enabled Then
    SnStr = HexPad(cpu.ShiftRight(Std(1).eax, 16), 4) & "-" & _
            HexPad(Std(1).eax And 65535, 4) & "-" & _
            HexPad(cpu.ShiftRight(Std(3).eax, 16), 4) & "-" & _
            HexPad(Std(3).edx And 65535, 4) & "-" & _
            HexPad(cpu.ShiftRight(Std(3).ecx, 16), 4) & "-" & _
            HexPad(Std(3).ecx And 65535, 4)
  End If
End Property

'Return whether CPUID is supported
Public Property Get Supported() As Boolean
  Supported = m_nLevelsStd <> 0
End Property

'Return the vendor
Public Property Get Vendor() As eVendor
  If m_nLevelsStd > 0 Then
    Vendor = m_nVendor
  End If
End Property

'Return the Vendor ID string
Public Property Get VendorIdStr(Optional Standard As Boolean = True) As String
  If Standard Then
    VendorIdStr = DecodeValue(Std(0).ebx) & _
                  DecodeValue(Std(0).edx) & _
                  DecodeValue(Std(0).ecx)
  Else
    VendorIdStr = DecodeValue(Ext(0).ebx) & _
                  DecodeValue(Ext(0).edx) & _
                  DecodeValue(Ext(0).ecx)
  End If
End Property

'--Extended register methods

'Extended: number of bits of physical address supported
Public Property Get xAddrBitsPhysical() As Long
  If m_nLevelsExt >= 8 Then
    xAddrBitsPhysical = Ext(8).eax And &HFF
  End If
End Property

'Extended: number of bits of virtual address supported
Public Property Get xAddrBitsVirtual() As Long
  If m_nLevelsExt >= 8 Then
    xAddrBitsVirtual = cpu.ShiftRight(Ext(8).eax, 8) And &HFF
  End If
End Property

'Extended: return a cache field value
Public Function xCache(Field As eExCacheFields) As Long
  If m_nLevelsExt >= 6 Then
    Select Case Field
    Case eExCacheFields.ecf_L1_4MbDataAssoc:        xCache = cpu.ShiftRight(Ext(5).eax, 24)
    Case eExCacheFields.ecf_L1_4MbDataEntries:      xCache = cpu.ShiftRight(Ext(5).eax, 16) And &HFF
    Case eExCacheFields.ecf_L1_4MbCodeAssoc:        xCache = cpu.ShiftRight(Ext(5).eax, 8) And &HFF
    Case eExCacheFields.ecf_L1_4MbCodeEntries:      xCache = Ext(5).eax And &HFF
    
    Case eExCacheFields.ecf_L1_4KbDataAssoc:        xCache = cpu.ShiftRight(Ext(5).ebx, 24)
    Case eExCacheFields.ecf_L1_4KbDataEntries:      xCache = cpu.ShiftRight(Ext(5).ebx, 16) And &HFF
    Case eExCacheFields.ecf_L1_4KbCodeAssoc:        xCache = cpu.ShiftRight(Ext(5).ebx, 8) And &HFF
    Case eExCacheFields.ecf_L1_4KbCodeEntries:      xCache = Ext(5).ebx And &HFF
    
    Case eExCacheFields.ecf_L1_DataSizeKb:          xCache = cpu.ShiftRight(Ext(5).ecx, 24)
    Case eExCacheFields.ecf_L1_DataAssoc:           xCache = cpu.ShiftRight(Ext(5).ecx, 16) And &HFF
    Case eExCacheFields.ecf_L1_DataLinesPerTag:     xCache = cpu.ShiftRight(Ext(5).ecx, 8) And &HFF
    Case eExCacheFields.ecf_L1_DataLineSizeBytes:   xCache = Ext(5).ecx And &HFF
    
    Case eExCacheFields.ecf_L1_CodeSizeKb:          xCache = cpu.ShiftRight(Ext(5).edx, 24)
    Case eExCacheFields.ecf_L1_CodeAssoc:           xCache = cpu.ShiftRight(Ext(5).edx, 16) And &HFF
    Case eExCacheFields.ecf_L1_CodeLinesPerTag:     xCache = cpu.ShiftRight(Ext(5).edx, 8) And &HFF
    Case eExCacheFields.ecf_L1_CodeLineSizeBytes:   xCache = Ext(5).edx And &HFF
    
    Case eExCacheFields.ecf_L2_4MbDataAssoc:        xCache = cpu.ShiftRight(Ext(6).eax, 28)
    Case eExCacheFields.ecf_L2_4MbDataEntries:      xCache = cpu.ShiftRight(Ext(6).eax, 16) And &HFFF
    Case eExCacheFields.ecf_L2_4MbCodeAssoc:        xCache = cpu.ShiftRight(Ext(6).eax, 12) And &HF
    Case eExCacheFields.ecf_L2_4MbCodeEntries:      xCache = Ext(6).eax And &HFFF
    
    Case eExCacheFields.ecf_L2_4KbDataAssoc:        xCache = cpu.ShiftRight(Ext(6).ebx, 28)
    Case eExCacheFields.ecf_L2_4KbDataEntries:      xCache = cpu.ShiftRight(Ext(6).ebx, 16) And &HFFF
    Case eExCacheFields.ecf_L2_4KbCodeAssoc:        xCache = cpu.ShiftRight(Ext(6).ebx, 12) And &HF
    Case eExCacheFields.ecf_L2_4KbCodeEntries:      xCache = Ext(6).ebx And &HFFF
    
    Case eExCacheFields.ecf_L2_CacheSizeKb:         xCache = cpu.ShiftRight(Ext(6).ecx, 16)
    Case eExCacheFields.ecf_L2_CacheAssoc:          xCache = cpu.ShiftRight(Ext(6).ecx, 12) And &HF
    Case eExCacheFields.ecf_L2_CacheLinesPerTag:    xCache = cpu.ShiftRight(Ext(6).ecx, 8) And &HF
    Case eExCacheFields.ecf_L2_CacheLineSizeBytes:  xCache = Ext(6).ecx And &HFF
    End Select
  End If
End Function

'Extended: return an extended cache association value as a string
Public Function xCacheAssocStr(nType As Long) As String
  If m_nLevelsExt >= 6 Then
    Select Case nType
    Case &H0:   xCacheAssocStr = "Off"
    Case &H1:   xCacheAssocStr = "Direct mapped"
    Case &H2:   xCacheAssocStr = "2-way"
    Case &H4:   xCacheAssocStr = "4-way"
    Case &H6:   xCacheAssocStr = "8-way"
    Case &H8:   xCacheAssocStr = "16-way"
    Case &HFF:  xCacheAssocStr = "Full"
    End Select
  End If
End Function

'Extended: power management support
Public Function xPowerManagement(nType As eExPowerManagment) As Boolean
  If m_nLevelsExt >= 7 Then
    If nType >= 0 Then
      If nType <= 31 Then
        xPowerManagement = cpu.ShiftRight(Ext(7).edx, nType) And &H1
      End If
    End If
  End If
End Function

'--Private routines
Private Function DecodeValue(nValue As Long) As String
  Dim i As Long
  
  For i = 0 To 24 Step 8
    DecodeValue = DecodeValue & Chr$(cpu.ShiftRight(nValue, i) And &HFF)
  Next i
End Function
