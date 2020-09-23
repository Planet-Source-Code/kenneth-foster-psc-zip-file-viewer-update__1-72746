VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form Form1 
   Caption         =   "PSC Zip Viewer"
   ClientHeight    =   8925
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14655
   LinkTopic       =   "Form1"
   ScaleHeight     =   8925
   ScaleWidth      =   14655
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rtb1 
      Height          =   8730
      Left            =   9135
      TabIndex        =   17
      Top             =   150
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   15399
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"Form1.frx":0000
   End
   Begin VB.Frame Frame2 
      Caption         =   "Help"
      Height          =   2820
      Left            =   4800
      TabIndex        =   11
      Top             =   555
      Visible         =   0   'False
      Width           =   3795
      Begin VB.CommandButton cmdCloseHelp 
         Caption         =   "Exit"
         Height          =   285
         Left            =   2970
         TabIndex        =   12
         Top             =   2460
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   $"Form1.frx":0082
         Height          =   615
         Left            =   270
         TabIndex        =   16
         Top             =   1935
         Width           =   3480
      End
      Begin VB.Label Label5 
         Caption         =   "3. Import Zip File lets you open a zip file not on the      list. It does not place zip file in folder."
         Height          =   450
         Left            =   105
         TabIndex        =   15
         Top             =   1380
         Width           =   3540
      End
      Begin VB.Label Label4 
         Caption         =   $"Form1.frx":0112
         Height          =   630
         Left            =   105
         TabIndex        =   14
         Top             =   750
         Width           =   3510
      End
      Begin VB.Label Label3 
         Caption         =   "1. All associated file names MUST be the same as     the zip file."
         Height          =   390
         Left            =   105
         TabIndex        =   13
         Top             =   285
         Width           =   3615
      End
   End
   Begin VB.TextBox txtAuthorsName 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4410
      MultiLine       =   -1  'True
      TabIndex        =   3
      Text            =   "Form1.frx":019D
      Top             =   3735
      Width           =   4590
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   4710
      Left            =   4290
      TabIndex        =   2
      Top             =   4125
      Width           =   4815
   End
   Begin VB.PictureBox pic1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3555
      Left            =   4410
      ScaleHeight     =   235
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   304
      TabIndex        =   1
      Top             =   150
      Width           =   4590
      Begin VB.Frame Frame1 
         Caption         =   "Create Author Text File"
         Height          =   3030
         Left            =   135
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   4305
         Begin VB.CommandButton cmdClose 
            Caption         =   "Exit"
            Height          =   360
            Left            =   3090
            TabIndex        =   10
            Top             =   2010
            Width           =   810
         End
         Begin VB.CommandButton cmdSaveName 
            Caption         =   "Save"
            Height          =   345
            Left            =   675
            TabIndex        =   9
            Top             =   2085
            Width           =   750
         End
         Begin VB.TextBox txtSaveAs 
            Height          =   345
            Left            =   540
            TabIndex        =   6
            Top             =   1425
            Width           =   3345
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   510
            TabIndex        =   5
            Top             =   675
            Width           =   3360
         End
         Begin VB.Label Label2 
            Caption         =   "Save As:"
            Height          =   195
            Left            =   540
            TabIndex        =   8
            Top             =   1185
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Name:"
            Height          =   180
            Left            =   540
            TabIndex        =   7
            Top             =   435
            Width           =   885
         End
      End
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8670
      Left            =   15
      TabIndex        =   0
      Top             =   150
      Width           =   4245
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import Zip File"
      End
      Begin VB.Menu mnuTextFileAuthor 
         Caption         =   "Create Author Text File"
      End
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu dash1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Project by Ken Foster 2009 Dec
'credits to Rde and Clint LaFever for their code (see modules)

Option Explicit

Private bzip As CGUnzipFiles
Dim sBuffer As String

Private Sub Form_Load()
    Set bzip = New CGUnzipFiles
    InitKeyWords
    File1.Path = App.Path & "\ZippedFiles\"
    sBuffer = "C:\"
End Sub
    
Private Sub Form_Unload(Cancel As Integer)
    KillFolder App.Path & "\tempUnzip\", True  'dump any files that are still in tempUnzip folder
    DoEvents
    Set bzip = Nothing
    Unload Me
End Sub
    
Private Sub List1_Click()
    'put code in textbox
    rtb1.Text = ""
    rtb1.Text = FileText(List1.Text)
    ColorIn rtb1
End Sub
    
Private Sub File1_Click()
    Dim FNlgth As Integer
    Dim FNstrg As String
    Dim ext As String
    
    KillFolder App.Path & "\tempUnzip\", True      'remove any old files
    If File1.FileName = "" Then Exit Sub
    FNlgth = Len(File1.FileName)                   'remove extension
    FNstrg = Left$(File1.FileName, FNlgth - 4)
    'put correct extension on picture file
    If Dir(App.Path & "\Photos\" & FNstrg & ".jpg") <> "" Then ext = ".jpg"
    If Dir(App.Path & "\Photos\" & FNstrg & ".jpeg") <> "" Then ext = ".jpeg"
    If Dir(App.Path & "\Photos\" & FNstrg & ".gif") <> "" Then ext = ".gif"
    If Dir(App.Path & "\Photos\" & FNstrg & ".bmp") <> "" Then ext = ".bmp"
    'if no picture
    If Dir(App.Path & "\Photos\" & FNstrg & ext) = "" Then GoTo skiptohere
    'load picture
    StretchSourcePictureFromFile App.Path & "\Photos\" & FNstrg & ext, pic1
    'load authors name into textbox
    txtAuthorsName.Text = "Authors Name: " & FileText(App.Path & "\AuthorFiles\" & FNstrg & ".txt")
    'unzip file
    bzip.Unzip App.Path & "\ZippedFiles\" & File1.FileName, App.Path & "\tempUnzip\"
    'clear so we can load new info
    List1.Clear
    rtb1.Text = ""
    'load list1 with files
    FileList App.Path & "\tempUnzip\"
    Exit Sub
skiptohere:
    pic1.Picture = LoadPicture()
    pic1.CurrentX = pic1.ScaleWidth / 2 - 70
    pic1.CurrentY = pic1.ScaleHeight / 2
    pic1.FontSize = 18
    pic1.FontBold = True
    pic1.Print "NO PICTURE"
End Sub
    
Private Function FileList(ByVal Pathname As String, Optional DirCount As Long, Optional FileCount As Long) As String
    Dim ShortName As String, LongName As String
    Dim NextDir As String
    Dim fnExt As String
    On Error Resume Next
    
    Static FolderList As Collection
    Set FolderList = Nothing                     'clear for next file
    Screen.MousePointer = vbHourglass
    If FolderList Is Nothing Then
        Set FolderList = New Collection
        FolderList.add Pathname
        DirCount = 0
        FileCount = 0
    End If
    
    Do
        NextDir = FolderList.Item(1)
        FolderList.Remove 1
        ShortName = Dir(NextDir & "\*.*", vbNormal Or vbArchive Or vbDirectory)
        
        Do While ShortName > ""
            
            If ShortName = "." Or ShortName = ".." Then
            Else
                LongName = NextDir & "\" & ShortName
                fnExt = LCase(Right$(LongName, 4))               'get extension and make lower case
                'skip un-nessary files
                If fnExt = ".jpg" Or fnExt = ".jpeg" _
                Or fnExt = ".gif" Or fnExt = ".bmp" _
                Or fnExt = ".ico" Or fnExt = ".frx" _
                Or fnExt = ".scc" Or fnExt = ".ctx" _
                Then GoTo here
                'put files in listbox
                List1.AddItem LongName
here:
                If (GetAttr(LongName) And vbDirectory) > 0 Then
                    FolderList.add LongName
                    DirCount = DirCount + 1
                Else
                    FileList = FileList & LongName & vbCrLf
                    FileCount = FileCount + 1
                End If
            End If
            ShortName = Dir()
        Loop
    Loop Until FolderList.Count = 0
    Screen.MousePointer = vbNormal
End Function
    
Private Function FileText(ByVal FileName As String) As String
    Dim Handle As Integer
    On Error Resume Next
    
    If Len(Dir$(FileName)) = 0 Then
        Err.Raise 53
    End If
    
    Handle = FreeFile
    Open FileName$ For Binary As #Handle
    FileText = Space$(LOF(Handle))
    Get #Handle, , FileText
    Close #Handle
End Function
    
Private Sub mnuExit_Click()
   KillFolder App.Path & "\tempUnzip\", True  'dump any files that are still there
   DoEvents
   Set bzip = Nothing
   Unload Me
End Sub

Private Sub mnuHelp_Click()
   Frame2.Visible = True
End Sub

Private Sub mnuImport_Click()
    List1.Clear
    rtb1.Text = ""
    pic1.Picture = LoadPicture()
    File1.ListIndex = -1
    
    sBuffer = GetFolder(Me.hWnd, sBuffer, "Open a File", True, False)
    If sBuffer = "" Then Exit Sub
    bzip.Unzip sBuffer, App.Path & "\tempUnzip\"
    'load list1 with files
    FileList App.Path & "\tempUnzip\"
End Sub

Private Sub mnuTextFileAuthor_Click()
   Frame1.Visible = True
End Sub

Private Sub cmdClose_Click()
   Frame1.Visible = False
End Sub

Private Sub cmdCloseHelp_Click()
   Frame2.Visible = False
End Sub

Private Sub cmdSaveName_Click()
On Error GoTo Handle
Dim sTemp As String
Dim filenum As Integer
Dim rply As String

    filenum = FreeFile
    sTemp = txtName.Text

    If FileExists(App.Path & "\AuthorFiles\" & txtSaveAs.Text & ".txt") = False Then    'Check whether the file created
        Open App.Path & "\AuthorFiles\" & txtSaveAs.Text & ".txt" For Output As #filenum  'Opening the file to SaveText
        Print #filenum, sTemp             'Printing  the text to the file
        Close #filenum                        'Closing
    Else
        rply = MsgBox("File already exists. Do you want to overwrite?", vbYesNo, "File exists already")
        If rply = vbYes Then
           Open App.Path & "\AuthorFiles\" & txtSaveAs.Text & ".txt" For Output As #filenum  'Opening the file to SaveText
           Print #filenum, sTemp             'Printing  the text to the file
           Close #filenum                        'Closing
        Else
           Open App.Path & "\AuthorFiles\" & txtSaveAs.Text & ".txt" For Append As #filenum  'Opening the file to SaveText
           Print #filenum, sTemp             'Printing  the text to the file
           Close #filenum                        'Closing
       End If
    End If
    MsgBox "File Saved " & App.Path & "\AuthorFiles\" & txtSaveAs.Text & ".txt"
Exit Sub
Handle:
    
    MsgBox "Error " & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"

End Sub

Private Function FileExists(FileName As String) As Boolean
'This function checks the existance of a file
On Error GoTo Handle
    If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
Handle:
    FileExists = False
End Function

