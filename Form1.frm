VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   1830
   ClientTop       =   2655
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6585
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2400
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "E&xtract"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Delete"
      Height          =   255
      Left            =   2880
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   4560
      Pattern         =   "*.jpg"
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "File Size :"
      Height          =   255
      Left            =   3960
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Private Type Archive
    strFileName As String
    lngSize As Long
    lngOffset As Long
End Type

Private udtArchive() As Archive

Private Sub Combo1_Click()
    With Combo1
        Label1.Caption = "File Size : " & GetFile(.ItemData(.ListIndex)) & " bytes"
    End With
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    With Combo1
        Call DeleteFile(.ItemData(.ListIndex))
        Call GetFileNames
    End With
    If Combo1.ListCount > 0 Then Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    With CommonDialog1
        .Flags = cdlOFNFileMustExist + cdlOFNExplorer
        .Filter = "All Files|*.*"
        .ShowOpen
        Call AddFile(.FileName)
        Call GetFileNames
    End With
End Sub

Private Sub Command3_Click()
    With CommonDialog1
        .Flags = cdlOFNCreatePrompt + cdlOFNExplorer + cdlOFNOverwritePrompt + cdlOFNPathMustExist
        .DefaultExt = Mid$(Combo1.Text, InStrRev(Combo1.Text, ".") + 1)
        .Filter = "*." & .DefaultExt & "||All Files|*.*"
        .FileName = Combo1.Text
        .ShowSave
        Call ExtractFile(Combo1.ItemData(Combo1.ListIndex), .FileName)
        Call GetFileNames
    End With
End Sub

Private Sub Form_Load()
    'Call ArchiveFiles
    Call GetFileNames
End Sub

Private Sub GetFileNames()
    Dim strSize As String
    Dim lngOffset As Long
    Dim lngFileLen As Long
    
    If Len(Dir(App.Path & "\Archive.fsf")) > 0 Then
        lngFileLen = FileLen(App.Path & "\Archive.fsf")
        Open App.Path & "\Archive.fsf" For Binary Access Read As #1
            Erase udtArchive
            Combo1.Clear
            Label1.Caption = ""
            Set Image1.Picture = Nothing
            Do While Not EOF(1) And lngOffset < LOF(1)
                On Error Resume Next
                ReDim Preserve udtArchive(UBound(udtArchive) + 1)
                If Err.Number <> 0 Then
                    ReDim udtArchive(0)
                    Err.Clear
                End If
                On Error GoTo 0
                
                With udtArchive(UBound(udtArchive))
                    .strFileName = String(255, Chr$(0))
                    strSize = String(255, Chr$(0))
                    Get #1, lngOffset + 1, .strFileName
                    Get #1, , strSize
                    .lngSize = Val(strSize)
                    .strFileName = Left$(.strFileName, InStr(1, .strFileName, Chr$(0), vbTextCompare) - 1)
                    If Len(.strFileName) > 0 Then
                        lngOffset = lngOffset + 255 + Len(strSize) + 1
                        .lngOffset = lngOffset
                        With Combo1
                            .AddItem udtArchive(UBound(udtArchive)).strFileName
                            .ItemData(.NewIndex) = UBound(udtArchive)
                        End With
                        lngOffset = lngOffset + .lngSize - 1
                    End If
                End With
            Loop
        Close #1
        If Combo1.ListCount > 0 Then Combo1.ListIndex = 0
    End If
End Sub

Private Function GetFile(lngIndex As Long) As Long
    Dim picArray() As Byte
    Dim strFile As String
    
    GetFile = 0
    If Len(Dir(App.Path & "\Archive.fsf")) > 0 Then
        Open App.Path & "\Archive.fsf" For Binary Access Read As #1
            With udtArchive(lngIndex)
                strFile = String(.lngSize, Chr$(0))
                Get #1, .lngOffset, strFile
                picArray = StrConv(strFile, vbFromUnicode)
                Set Image1.Picture = PictureFromByteStream(picArray)
                GetFile = .lngSize
            End With
        Close #1
    End If
End Function

Private Sub ArchiveFiles()
    Dim lngCounter As Long
    Dim strSize As String
    Dim strFile As String
    Dim strFileName As String
    
    On Error Resume Next
    Kill App.Path & "\Archive.fsf"
    On Error GoTo 0
    
    File1.Path = App.Path
    Open App.Path & "\Archive.fsf" For Binary Access Write As #1
        For lngCounter = 0 To File1.ListCount - 1
            strFile = String(FileLen(App.Path & "\" & File1.List(lngCounter)), Chr$(0))
            Open App.Path & "\" & File1.List(lngCounter) For Binary Access Read As #2
                Get #2, 1, strFile
            Close #2
            strFileName = File1.List(lngCounter)
            strFileName = strFileName & String(255 - Len(strFileName), Chr$(0))
            strSize = CStr(FileLen(App.Path & "\" & File1.List(lngCounter)))
            strSize = strSize & String(255 - Len(strSize), Chr$(0))
            Put #1, , strFileName
            Put #1, , strSize
            Put #1, , strFile
        Next
    Close #1
End Sub

Private Sub DeleteFile(lngIndex As Long)
    Dim lngCounter As Long
    Dim strData As String
    Dim strFile As String
    Dim strTempFile As String
    Dim strTempDirectory As String
    
    If Len(Dir(App.Path & "\Archive.fsf")) > 0 Then
        strTempDirectory = String(255, Chr$(0))
        GetTempPath 255, strTempDirectory
        strTempDirectory = Left$(strTempDirectory, InStr(1, strTempDirectory, Chr$(0)) - 1)
        
        strTempFile = String(255, 0)
        GetTempFileName strTempDirectory, "FSF", 0, strTempFile
        strTempFile = Left$(strTempFile, InStr(1, strTempFile, Chr$(0)) - 1)
                
        Open App.Path & "\Archive.fsf" For Binary Access Read As #1
        Open strTempFile For Binary Access Write As #2
        
        For lngCounter = 0 To UBound(udtArchive)
            With udtArchive(lngCounter)
                strData = String(510, Chr$(0))
                Get #1, , strData
                strFile = String(.lngSize, Chr$(0))
                Get #1, , strFile
                If lngCounter <> lngIndex Then
                    Put #2, , strData
                    Put #2, , strFile
                End If
            End With
        Next
        Close #1
        Close #2
        Kill App.Path & "\Archive.fsf"
        FileCopy strTempFile, App.Path & "\Archive.fsf"
        Kill strTempFile
    End If
End Sub

Private Sub AddFile(pFileName As String)
    Dim strSize As String
    Dim strFile As String
    Dim strFileName As String
    Dim lngOffset As Long
    
    If Len(pFileName) > 0 Then
        lngOffset = FileLen(App.Path & "\Archive.fsf")
        Open App.Path & "\Archive.fsf" For Binary Access Write As #3
            strFile = String(FileLen(pFileName), Chr$(0))
            Open pFileName For Binary Access Read As #4
                Get #4, 1, strFile
            Close #4
            strFileName = Mid$(pFileName, InStrRev(pFileName, "\") + 1)
            strFileName = strFileName & String(255 - Len(strFileName), Chr$(0))
            strSize = CStr(FileLen(pFileName))
            strSize = strSize & String(255 - Len(strSize), Chr$(0))
            Put #3, lngOffset + 1, strFileName
            Put #3, , strSize
            Put #3, , strFile
        Close #3
    End If
End Sub

Private Sub ExtractFile(lngIndex As Long, pFileName As String)
    Dim strFile As String
    
    If Len(pFileName) > 0 Then
        With udtArchive(lngIndex)
            Open App.Path & "\Archive.fsf" For Binary Access Read As #1
                strFile = String(.lngSize, Chr$(0))
                Get #1, .lngOffset, strFile
            Close #1
            
            On Error Resume Next
            Kill pFileName
            On Error GoTo 0
            
            Open pFileName For Binary Access Write As #2
                Put #2, 1, strFile
            Close #2
        End With
    End If
End Sub

Private Sub Form_Resize()
    With Image1
        .Move 30, Command1.Top + Command1.Height + 30, Me.ScaleWidth - 60, Me.ScaleHeight - .Top - 30
    End With
End Sub
