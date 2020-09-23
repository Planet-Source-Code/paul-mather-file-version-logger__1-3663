VERSION 5.00
Begin VB.Form IDD_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Logger"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4020
   Icon            =   "IDD_Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox IDCK_StaggerOutput 
      Caption         =   "Stagger Output"
      Height          =   255
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.TextBox IDE_SearchString 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "*.*"
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton IDCM_Exit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton IDCM_Execute 
      Caption         =   "&Execute"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox IDCK_Recursive 
      Caption         =   "Recurse Directories"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   2760
      Value           =   1  'Checked
      Width           =   1935
   End
   Begin VB.DirListBox IDDIR_Local 
      Height          =   1890
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.DriveListBox IDDRIVE_Local 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   3735
   End
   Begin VB.Label IDL_SearchString 
      AutoSize        =   -1  'True
      Caption         =   "Search String"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   960
   End
End
Attribute VB_Name = "IDD_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LevelsDeep As Integer

Private Sub Form_Load()
    LevelsDeep = 0
End Sub

Private Sub IDCM_Execute_Click()
Dim LocalPath As String
Dim LogFileName As String
Dim PrintString As String
Dim FileCount As Integer
Dim DirCount As Integer
    
    Me.MousePointer = vbHourglass
    If Right(IDDIR_Local.path, 1) = "\" Then
        LocalPath = IDDIR_Local.path
    Else
        LocalPath = IDDIR_Local.path & "\"
    End If

    LogFileName = Environ("TEMP") & "\" & App.Title & ".log"

    Open LogFileName For Output As #1
    PrintString = "File Log for Directory: " & Mid(LocalPath, 1, Len(LocalPath) - 1)
    If IDCK_Recursive.Value = vbChecked Then
        PrintString = PrintString & " (Recursing Sub Directories)"
    End If
    Print #1, PrintString
    Print #1, Now
    Print #1,
    
    Call FindFiles(LocalPath, IDE_SearchString, FileCount, DirCount, CInt(IDCK_Recursive.Value) * -1)
    
    Close #1
    
    Call Shell("notepad " & LogFileName, vbMaximizedFocus)
    Me.MousePointer = vbDefault
    
End Sub

Function FindFiles(path As String, SearchStr As String, FileCount As Integer, DirCount As Integer, Optional ByVal RecurseSubs As Boolean = True)
Dim FileInformation As FILE_INFORMATION
Dim FileName As String   ' Walking filename variable.
Dim DirName As String    ' SubDirectory Name.
Dim dirNames() As String ' Buffer for directory name entries.
Dim nDir As Integer      ' Number of directories in this path.
Dim i As Integer         ' For-loop counter.
Dim VersionInfo As String
Dim LastDir As String
Dim Spaces As String
Dim Count As Integer
Const SpaceChar As String = vbTab


    On Error GoTo sysFileERR
    If Right(path, 1) <> "\" Then path = path & "\"
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    DirName = Dir(path, vbDirectory Or vbHidden)  ' Even if hidden.
    Do While Len(DirName) > 0
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetAttr(path & DirName) And vbDirectory Then
               dirNames(nDir) = DirName
               DirCount = DirCount + 1
               nDir = nDir + 1
               ReDim Preserve dirNames(nDir)
            End If
sysFileERRCont:
        End If
        DirName = Dir()  ' Get next subdirectory.
    Loop
    
    ' Search through this directory and sum file sizes.
    FileName = Dir(path & SearchStr, vbNormal Or vbHidden Or vbSystem Or vbReadOnly)
    While Len(FileName) <> 0
        FindFiles = FindFiles + FileLen(path & FileName)
        FileCount = FileCount + 1
        Call GetFileInformation(path & FileName, FileInformation)
        Spaces = ""
        If IDCK_StaggerOutput.Value = vbChecked Then
            If LevelsDeep > 0 Then
                For Count = 1 To LevelsDeep
                    Spaces = Spaces & SpaceChar
                Next Count
            End If
        Else
            Spaces = ""
        End If
        If FileInformation.cDirectory <> LastDir Then
            Print #1,
            Print #1, Spaces & "Directory: --> " & FileInformation.cDirectory
            Print #1, Spaces & "-----------------------------------------------------------------------"
            LastDir = FileInformation.cDirectory
        End If
        If FileInformation.nVerMajor <> 0 Or FileInformation.nVerMinor <> 0 Or FileInformation.nVerRevision <> 0 Then
            VersionInfo = " - Version:" & FileInformation.nVerMajor & "." & FileInformation.nVerMinor & "." & FileInformation.nVerRevision
        Else
            VersionInfo = ""
        End If
'        Spaces = Spaces & SpaceChar
        Print #1, Spaces & FileInformation.cFilename & " - Modify Date:" & Format(FileInformation.dtLastModifyTime, "mm/dd/yyyy HH:MM AMPM") & " - File Size:" & FileInformation.nFileSize & " bytes" & VersionInfo
        FileName = Dir()  ' Get next file.
    Wend

    ' If there are sub-directories..
    If nDir > 0 And RecurseSubs = True Then
        ' Recursively walk into them
        For i = 0 To nDir - 1
            LevelsDeep = LevelsDeep + 1
            FindFiles = FindFiles + FindFiles(path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
            LevelsDeep = LevelsDeep - 1
        Next i
    End If

AbortFunction:
    Exit Function
sysFileERR:
    If Right(DirName, 4) = ".sys" Then
        Resume sysFileERRCont ' Known issue with pagefile.sys
    Else
        Resume AbortFunction
    End If
End Function


Private Sub IDCM_Exit_Click()
    Unload Me
End Sub
Private Sub IDDRIVE_Local_Change()
    On Error Resume Next
    IDDIR_Local.path = IDDRIVE_Local.Drive
End Sub
