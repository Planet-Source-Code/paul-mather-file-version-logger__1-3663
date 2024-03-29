Attribute VB_Name = "IDBAS_FileVersion"
Option Explicit

Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILE_TIME, lpLastAccessTime As FILE_TIME, lpLastWriteTime As FILE_TIME) As Long
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OF_STRUCT, ByVal wStyle As Long) As Long
Private Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILE_TIME, lpLocalFileTime As FILE_TIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILE_TIME, lpSystemTime As SYSTEM_TIME) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, ByVal Source As Long, ByVal Length As Long)

Private Const OF_READ = &H0
Private Const OF_SHARE_DENY_NONE = &H40
Private Const OFS_MAXPATHNAME = 128

' ===== From Win32 Ver.h =================
' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FFI_SIGNATURE = &HFEEF04BD
Private Const VS_FFI_STRUCVERSION = &H10000
Private Const VS_FFI_FILEFLAGSMASK = &H3F&

' ----- VS_VERSION.dwFileFlags -----
Private Const VS_FF_DEBUG = &H1
Private Const VS_FF_PRERELEASE = &H2
Private Const VS_FF_PATCHED = &H4
Private Const VS_FF_PRIVATEBUILD = &H8
Private Const VS_FF_INFOINFERRED = &H10
Private Const VS_FF_SPECIALBUILD = &H20

' ----- VS_VERSION.dwFileOS -----
Private Const VOS_UNKNOWN = &H0
Private Const VOS_DOS = &H10000
Private Const VOS_OS216 = &H20000
Private Const VOS_OS232 = &H30000
Private Const VOS_NT = &H40000
Private Const VOS__BASE = &H0
Private Const VOS__WINDOWS16 = &H1
Private Const VOS__PM16 = &H2
Private Const VOS__PM32 = &H3
Private Const VOS__WINDOWS32 = &H4

Private Const VOS_DOS_WINDOWS16 = &H10001
Private Const VOS_DOS_WINDOWS32 = &H10004
Private Const VOS_OS216_PM16 = &H20002
Private Const VOS_OS232_PM32 = &H30003
Private Const VOS_NT_WINDOWS32 = &H40004


' ----- VS_VERSION.dwFileType -----
Private Const VFT_UNKNOWN = &H0
Private Const VFT_APP = &H1
Private Const VFT_DLL = &H2
Private Const VFT_DRV = &H3
Private Const VFT_FONT = &H4
Private Const VFT_VXD = &H5
Private Const VFT_STATIC_LIB = &H7

' ----- VS_VERSION.dwFileSubtype for VFT_WINDOWS_DRV -----
Private Const VFT2_UNKNOWN = &H0
Private Const VFT2_DRV_PRINTER = &H1
Private Const VFT2_DRV_KEYBOARD = &H2
Private Const VFT2_DRV_LANGUAGE = &H3
Private Const VFT2_DRV_DISPLAY = &H4
Private Const VFT2_DRV_MOUSE = &H5
Private Const VFT2_DRV_NETWORK = &H6
Private Const VFT2_DRV_SYSTEM = &H7
Private Const VFT2_DRV_INSTALLABLE = &H8
Private Const VFT2_DRV_SOUND = &H9
Private Const VFT2_DRV_COMM = &HA

Private Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer ' e.g. = &h0000 = 0
    dwStrucVersionh As Integer ' e.g. = &h0042 = .42
    dwFileVersionMSl As Integer ' e.g. = &h0003 = 3
    dwFileVersionMSh As Integer ' e.g. = &h0075 = .75
    dwFileVersionLSl As Integer ' e.g. = &h0000 = 0
    dwFileVersionLSh As Integer ' e.g. = &h0031 = .31
    dwProductVersionMSl As Integer ' e.g. = &h0003 = 3
    dwProductVersionMSh As Integer ' e.g. = &h0010 = .1
    dwProductVersionLSl As Integer ' e.g. = &h0000 = 0
    dwProductVersionLSh As Integer ' e.g. = &h0031 = .31
    dwFileFlagsMask As Long ' = &h3F For version "0.42"
    dwFileFlags As Long ' e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long ' e.g. VOS_DOS_WINDOWS16
    dwFileType As Long ' e.g. VFT_DRIVER
    dwFileSubtype As Long ' e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long ' e.g. 0
    dwFileDateLS As Long ' e.g. 0
End Type

Public Type FILE_ATTRIBUTES
    bArchive As Boolean
    bCompressed As Boolean
    bDirectory As Boolean
    bHidden As Boolean
    bNormal As Boolean
    bReadOnly As Boolean
    bSystem As Boolean
    bTemporary As Boolean
End Type

Public Type FILE_INFORMATION
    cFilename As String
    cDirectory As String
    cFullFilePath As String
    cFileType As String
    nVerMajor As Long
    nVerMinor As Long
    nVerRevision As Long
    nFileSize As Long
    nFileAttributes As Long
    nFileType As Long
    faFileAttributes As FILE_ATTRIBUTES
    dtCreationDate As Date
    dtLastModifyTime As Date
    dtLastAccessTime As Date
End Type

Private Type SYSTEM_TIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type FILE_TIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type OF_STRUCT
     cBytes As Byte
     fFixedDisk As Byte
     nErrCode As Integer
     Reserved1 As Integer
     Reserved2 As Integer
     szPathName(OFS_MAXPATHNAME) As Byte
End Type

Public Function GetFileInformation(ByVal fileFullPath As String, ByRef FileInformation As FILE_INFORMATION, Optional ByVal showMsgBox As Boolean = False) As Boolean
Dim lDummy As Long, lsize As Long, rc As Long
Dim lVerbufferLen As Long, lVerPointer As Long
Dim sBuffer() As Byte
Dim udtVerBuffer As VS_FIXEDFILEINFO
Dim hFile As Integer
Dim FileStruct As OF_STRUCT
Dim CreationTime As FILE_TIME
Dim LastAccessTime As FILE_TIME
Dim LastWriteTime As FILE_TIME
Dim LocalFileTime As SYSTEM_TIME
Dim MessageString As String

    On Error GoTo e_HandleFileInformationError
    lsize = GetFileVersionInfoSize(fileFullPath, lDummy)
    If lsize >= 1 Then
        ReDim sBuffer(lsize)
        rc = GetFileVersionInfo(fileFullPath, 0&, lsize, sBuffer(0))
        rc = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerbufferLen)
        MoveMemory udtVerBuffer, lVerPointer, Len(udtVerBuffer)
    End If
    
    '**** Determine Filename Info ****
    FileInformation.cFullFilePath = fileFullPath
    FileInformation.cFilename = DetermineFilename(fileFullPath)
    FileInformation.cDirectory = DetermineDirectory(fileFullPath)
    
    '**** Determine File Date Info ****
    hFile = OpenFile(fileFullPath, FileStruct, OF_READ Or OF_SHARE_DENY_NONE)
    If GetFileTime(hFile, CreationTime, LastAccessTime, LastWriteTime) Then
        Call FileTimeToLocalFileTime(LastAccessTime, LastAccessTime)
        If Not FileTimeToSystemTime(LastAccessTime, LocalFileTime) Then
            FileInformation.dtLastAccessTime = Format(LocalFileTime.wMonth, "00") & "/" & Format(LocalFileTime.wDay, "00") & "/" & Format(LocalFileTime.wYear, "0000") & " " & Format(LocalFileTime.wHour, "00") & ":" & Format(LocalFileTime.wMinute, "00") & ":" & Format(LocalFileTime.wSecond, "00")
        End If
        Call FileTimeToLocalFileTime(CreationTime, CreationTime)
        If Not FileTimeToSystemTime(CreationTime, LocalFileTime) Then
            FileInformation.dtCreationDate = Format(LocalFileTime.wMonth, "00") & "/" & Format(LocalFileTime.wDay, "00") & "/" & Format(LocalFileTime.wYear, "0000") & " " & Format(LocalFileTime.wHour, "00") & ":" & Format(LocalFileTime.wMinute, "00") & ":" & Format(LocalFileTime.wSecond, "00")
        End If
        Call FileTimeToLocalFileTime(LastWriteTime, LastWriteTime)
        If Not FileTimeToSystemTime(LastWriteTime, LocalFileTime) Then
            FileInformation.dtLastModifyTime = Format(LocalFileTime.wMonth, "00") & "/" & Format(LocalFileTime.wDay, "00") & "/" & Format(LocalFileTime.wYear, "0000") & " " & Format(LocalFileTime.wHour, "00") & ":" & Format(LocalFileTime.wMinute, "00") & ":" & Format(LocalFileTime.wSecond, "00")
        End If
    End If

    Call lclose(hFile)

    '**** Determine File Attributes and Size
    FileInformation.nFileType = udtVerBuffer.dwFileType
    Select Case FileInformation.nFileType
        Case VFT_UNKNOWN
            FileInformation.cFileType = "Unknown"
        Case VFT_APP
            FileInformation.cFileType = "Application"
        Case VFT_DLL
            FileInformation.cFileType = "DLL Library"
        Case VFT_DRV
            FileInformation.cFileType = "Driver"
        Case VFT_FONT
            FileInformation.cFileType = "Font"
        Case VFT_VXD
            FileInformation.cFileType = "VXD File"
        Case VFT_STATIC_LIB
            FileInformation.cFileType = "Static Library"
        Case Else
            FileInformation.cFileType = "Unknown"
    End Select
    
    FileInformation.nFileAttributes = GetFileAttributes(fileFullPath)
    If FileInformation.nFileAttributes And &H20 Then
        FileInformation.faFileAttributes.bArchive = True
    Else
        FileInformation.faFileAttributes.bArchive = False
    End If
    If FileInformation.nFileAttributes And &H800 Then
        FileInformation.faFileAttributes.bCompressed = True
    Else
        FileInformation.faFileAttributes.bCompressed = False
    End If
    If FileInformation.nFileAttributes And &H10 Then
        FileInformation.faFileAttributes.bDirectory = True
    Else
        FileInformation.faFileAttributes.bDirectory = False
    End If
    If FileInformation.nFileAttributes And &H2 Then
        FileInformation.faFileAttributes.bHidden = True
    Else
        FileInformation.faFileAttributes.bHidden = False
    End If
    If FileInformation.nFileAttributes And &H80 Then
        FileInformation.faFileAttributes.bNormal = True
    Else
        FileInformation.faFileAttributes.bNormal = False
    End If
    If FileInformation.nFileAttributes And &H1 Then
        FileInformation.faFileAttributes.bReadOnly = True
    Else
        FileInformation.faFileAttributes.bReadOnly = False
    End If
    If FileInformation.nFileAttributes And &H4 Then
        FileInformation.faFileAttributes.bSystem = True
    Else
        FileInformation.faFileAttributes.bSystem = False
    End If
    If FileInformation.nFileAttributes And &H100 Then
        FileInformation.faFileAttributes.bTemporary = True
    Else
        FileInformation.faFileAttributes.bTemporary = False
    End If

    FileInformation.nFileSize = FileLen(fileFullPath)
    
    '**** Determine Product Version number ****
    If lsize >= 1 Then
        FileInformation.nVerMajor = udtVerBuffer.dwProductVersionMSh
        FileInformation.nVerMinor = udtVerBuffer.dwProductVersionMSl
        FileInformation.nVerRevision = udtVerBuffer.dwFileVersionLSl
    End If
    
    If showMsgBox = True Then
        MessageString = "Path:" & vbCr & "Filename:" & vbTab & vbTab & FileInformation.cFilename & vbCr & _
        "Directory:" & vbTab & vbTab & FileInformation.cDirectory & vbCr & _
        "Full Path:" & vbTab & vbTab & FileInformation.cFullFilePath & vbCr & vbCr & "Date:" & vbCr & _
        "Creation Date:" & vbTab & Format(FileInformation.dtCreationDate, "dddd, mmm dd yyyy H:MM:SS AMPM") & vbCr & _
        "Modify Date:" & vbTab & Format(FileInformation.dtLastModifyTime, "dddd, mmm dd yyyy H:MM:SS AMPM") & vbCr & _
        "Access Date:" & vbTab & Format(FileInformation.dtLastAccessTime, "dddd, mmm dd yyyy") & vbCr & vbCr & "Attributes:" & vbCr & _
        "Archive:" & vbTab & vbTab & FileInformation.faFileAttributes.bArchive & vbCr & _
        "Compressed:" & vbTab & FileInformation.faFileAttributes.bCompressed & vbCr & _
        "Directory:" & vbTab & vbTab & FileInformation.faFileAttributes.bDirectory & vbCr & _
        "Hidden:" & vbTab & vbTab & FileInformation.faFileAttributes.bHidden & vbCr & _
        "Normal:" & vbTab & vbTab & FileInformation.faFileAttributes.bNormal & vbCr & _
        "Read Only:" & vbTab & FileInformation.faFileAttributes.bReadOnly & vbCr & _
        "System:" & vbTab & vbTab & FileInformation.faFileAttributes.bSystem & vbCr & _
        "Temporary:" & vbTab & FileInformation.faFileAttributes.bTemporary & vbCr & vbCr & "Misc.:" & vbCr & _
        "File Size:" & vbTab & vbTab & Format(FileInformation.nFileSize / 1024, "###,###,### KB (") & Format(FileInformation.nFileSize, "###,###,### bytes)") & vbCr
        If FileInformation.nFileType <> VFT_UNKNOWN Then
            MessageString = MessageString & "File Type:" & vbTab & vbTab & FileInformation.cFileType & vbCr
        End If
        If lsize >= 1 Then
            MessageString = MessageString & "Version:" & vbTab & vbTab & FileInformation.nVerMajor & "." & FileInformation.nVerMinor & "." & FileInformation.nVerRevision
        End If
        Call MsgBox(MessageString, vbOKOnly + vbInformation, "Information")
    End If

    GetFileInformation = True
    Exit Function
    
e_HandleFileInformationError:
    GetFileInformation = False
    Exit Function
End Function

Private Function DetermineDirectory(inputString As String) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineDirectory = Mid(inputString, 1, pos)
End Function
Private Function DetermineFilename(inputString As String) As String
Dim pos As Integer
    pos = InStrRev(inputString, "\", , vbTextCompare)
    DetermineFilename = Mid(inputString, pos + 1, Len(inputString) - pos)
End Function
Private Function DetermineDrive(inputString As String) As String
Dim pos As Integer
    If inputString = "" Then Exit Function
    pos = InStr(1, inputString, ":\", vbTextCompare)
    DetermineDrive = Mid(inputString, 1, pos - 1)
End Function

