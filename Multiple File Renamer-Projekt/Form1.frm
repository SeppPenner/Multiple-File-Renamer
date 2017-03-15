VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple File Renamer"
   ClientHeight    =   6735
   ClientLeft      =   4815
   ClientTop       =   2925
   ClientWidth     =   10710
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":BD2A
   ScaleHeight     =   6735
   ScaleWidth      =   10710
   Begin VB.PictureBox PictureBox4 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8880
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   13
      Top             =   840
      Width           =   615
      Begin VB.Shape Shape2 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox PictureBox3 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8160
      ScaleHeight     =   375
      ScaleWidth      =   615
      TabIndex        =   12
      Top             =   840
      Width           =   615
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.PictureBox PictureBox2 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   8400
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      Top             =   120
      Width           =   500
   End
   Begin VB.PictureBox PictureBox1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   500
      Left            =   9000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      Top             =   120
      Width           =   500
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   7200
      TabIndex        =   9
      Top             =   3480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox NewName 
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   840
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9960
      Top             =   1800
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   4935
   End
   Begin VB.CommandButton SelectButton 
      Height          =   375
      Left            =   6360
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.DirListBox Dir1 
      Height          =   4815
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   4560
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   2040
      Width           =   4935
   End
   Begin VB.Label LabelName 
      Height          =   375
      Left            =   4560
      TabIndex        =   8
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   1050
      Left            =   9600
      Picture         =   "Form1.frx":17A54
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label LabelTime 
      Height          =   255
      Left            =   9600
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LanguageCaption 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   360
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ExtractIcon _
     Lib "shell32.dll" _
     Alias "ExtractIconA" _
     ( _
     ByVal hInst As Long, _
     ByVal lpszExeFileName As String, _
     ByVal nIconIndex As Long _
     ) _
   As Long
   
   Private Declare Function DrawIcon _
     Lib "user32" _
     ( _
     ByVal hdc As Long, _
     ByVal x As Long, _
     ByVal y As Long, _
     ByVal hIcon As Long _
     ) _
   As Long
   
   Private Declare Function DestroyIcon _
     Lib "user32" _
     ( _
     ByVal hIcon As Long _
     ) _
   As Long
   
   Private Declare Function ShellExecute Lib "shell32.dll" _
  Alias "ShellExecuteA" ( _
  ByVal hWnd As Long, _
  ByVal lpOperation As String, _
  ByVal lpFile As String, _
  ByVal lpParameters As String, _
  ByVal lpDirectory As String, _
  ByVal nShowCmd As Long) As Long

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dim d As Drive
    Set d = FSO.GetDrive(Left$(Drive1.Drive, 2))
    If d.IsReady Then
        Dir1.Path = Left$(Drive1.Drive, 1) & ":\"
    Else
        If Temporaer_Language = "GER" Then
            Message = MsgBox("Laufwerk nicht verfügbar", vbCritical + vbApplicationModal, "Multiple File Renamer")
            Drive1.Drive = FSO.GetDrive("C:")
        Else
            Message = MsgBox("Drive not available", vbCritical + vbApplicationModal, "Multiple File Renamer")
            Drive1.Drive = FSO.GetDrive("C:")
        End If
    End If
End Sub

Private Sub Form_Load()
    Call cmdExtractIcons(47, PictureBox1)
    Call cmdExtractIcons(23, PictureBox2)
    picShowPicture PictureBox3, (App.Path & "\German.bmp")
    picShowPicture PictureBox4, (App.Path & "\English.bmp")
    List1.Visible = False
    Count_Language_Fail = False
    ALTPressed = False
    Drive1.Drive = "C:"
    LanguagePath = App.Path & "\Language.ini"
    Set Inp = FSO.OpenTextFile(LanguagePath, ForReading)
    Select Case Inp.ReadLine
        Case "+GER"
            Language = "GER"
            Temporaer_Language = "GER"
            Call InitializeGerman
        Case "GER"
            Select Case Inp.ReadLine
                Case "+ENG"
                    Language = "ENG"
                    Temporaer_Language = "ENG"
                    Call InitializeEnglish
                Case "ENG"
                If Count_Language_Fail = False Then
                    Message = MsgBox("Keine Sprache ausgewählt!!/No language selected!!", vbCritical + vbApplicationModal, "Multiple File Renamer")
                    Count_Language_Fail = True
                    End
                End If
                End Select
    End Select
    Combo1.ListIndex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    Select Case Temporaer_Language
        Case "ENG"
            Select Case KeyCode
                Case vbKeyMenu
                    ALTPressed = True
                Case vbKeyR
                    If ALTPressed = True Then
                        Call SelectButton_Click
                        ALTPressed = False
                    End If
                Case vbKeyL
                    If ALTPressed = True Then
                        Temporaer_Language = "GER"
                        Call InitializeGerman
                        ALTPressed = False
                    End If
            End Select
        Case "GER"
            Select Case KeyCode
                Case vbKeyMenu
                        ALTPressed = True
                Case vbKeyU
                    If ALTPressed = True Then
                        Call SelectButton_Click
                        ALTPressed = False
                    End If
                Case vbKeyS
                    If ALTPressed = True Then
                        Temporaer_Language = "ENG"
                        Call InitializeEnglish
                        ALTPressed = False
                    End If
            End Select
    End Select
End Sub

Private Sub InitializeGerman()
    LanguageCaption.Caption = "Sprache: Deutsch"
    SelectButton.Caption = "Umbenennen (U)"
    LabelTime.Caption = "Zeit: " & Time
    LabelName.Caption = "Bitte neuen Dateinamen eingeben"
    Combo1.Clear
    Combo1.AddItem "Alle Dateien", 0
    Combo1.AddItem "Textdateien", 1
    Combo1.AddItem "Bilder", 2
    Combo1.AddItem "PDFs", 3
    Combo1.AddItem "Excel-Dateien", 4
    Combo1.AddItem "Word-Dateien", 5
    Combo1.AddItem "Powerpoint-Dateien", 6
    Combo1.AddItem "Videos", 7
    Combo1.AddItem "Musik", 8
    Combo1.ListIndex = Count_Listindex
End Sub

Private Sub InitializeEnglish()
    LanguageCaption.Caption = "Language: English"
    SelectButton.Caption = "Rename (R)"
    LabelTime.Caption = "Time: " & Time
    LabelName.Caption = "Please insert new file name"
    Combo1.Clear
    Combo1.AddItem "All files", 0
    Combo1.AddItem "Text files", 1
    Combo1.AddItem "Images", 2
    Combo1.AddItem "PDFs", 3
    Combo1.AddItem "Excel-Files", 4
    Combo1.AddItem "Word-Files", 5
    Combo1.AddItem "Powerpoint-Files", 6
    Combo1.AddItem "Videos", 7
    Combo1.AddItem "Music", 8
    Combo1.ListIndex = Count_Listindex
End Sub

Private Sub Image1_Click()
    URLGoTo Me.hWnd, "http://franzhuber23.blogspot.de/2014/07/multiple-file-renamer.html"
    On Error Resume Next
End Sub

Private Sub PictureBox1_Click() 'License
    Call Form2.Show
End Sub

Private Sub PictureBox2_Click() 'Developer
    Call Form3.Show
End Sub

Private Sub PictureBox3_Click() 'German
    If Temporaer_Language = "ENG" Then
        Temporaer_Language = "GER"
        Call InitializeGerman
    End If
End Sub

Private Sub PictureBox4_Click() 'English
    If Temporaer_Language = "GER" Then
        Temporaer_Language = "ENG"
        Call InitializeEnglish
    End If
End Sub

Private Sub SelectButton_Click()
    If NewName.Text <> "" Then
        Count_Files = 0
        Count_Help = 0
        List1.Clear
        SelectButton.Enabled = False
        NewName.Enabled = False
        Dim k As Long
        Dim s As String
        For k = 0 To File1.ListCount - 1
            If File1.Selected(k) Then
                s = File1.List(k)
                List1.AddItem s
                DoEvents
            End If
        Next
        If Right(File1.Path, 1) <> "\" Then
            PathWithoutFile = File1.Path + "\"
        Else
            PathWithoutFile = File1.Path
        End If
        If List1.ListCount > 0 Then
Above:      For k = 0 To List1.ListCount - 1
                Count_Help = Count_Help + 1
                If Count_Files <> 0 Then
                    If (Count_Files + 1) < 10 Then
                        DoEvents
                        PathWithFile = PathWithoutFile + List1.List(k)
                        FileEnding = Mid(PathWithFile, InStr(PathWithFile, "."))
                        TargetPath = PathWithoutFile + NewName.Text + "_" + CStr(0) + CStr(Count_Files + 1) + FileEnding
                        On Error GoTo Error_1
                        FSO.MoveFile PathWithFile, TargetPath
                        Count_Files = Count_Files + 1
                        DoEvents
                    Else
                        DoEvents
                        PathWithFile = PathWithoutFile + List1.List(k)
                        FileEnding = Mid(PathWithFile, InStr(PathWithFile, "."))
                        TargetPath = PathWithoutFile + NewName.Text + "_" + CStr(Count_Files + 1) + FileEnding
                        On Error GoTo Error_1
                        FSO.MoveFile PathWithFile, TargetPath
                        Count_Files = Count_Files + 1
                        DoEvents
                    End If
                Else
                    DoEvents
                    PathWithFile = PathWithoutFile + List1.List(k)
                    FileEnding = Mid(PathWithFile, InStr(PathWithFile, "."))
                    TargetPath = PathWithoutFile + NewName.Text + FileEnding
                    On Error GoTo Error_1
                    FSO.MoveFile PathWithFile, TargetPath
                    Count_Files = Count_Files + 1
                    DoEvents
                End If
            Next
            If Count_Help = List1.ListCount Then
                GoTo ENDE
            End If
        Else
            If Temporaer_Language = "GER" Then
                Message = MsgBox("Keine Dateien ausgewählt", vbInformation + vbApplicationModal, "Multiple File Renamer")
            Else
                Message = MsgBox("You haven't selected files", vbInformation + vbApplicationModal, "Multiple File Renamer")
            End If
        End If
    Else
        If Temporaer_Language = "GER" Then
            Message = MsgBox("Bitte neuen Namen für ihre Dateien eingeben", vbInformation + vbApplicationModal, "Multiple File Renamer")
        Else
            Message = MsgBox("Please insert a new name for your files", vbInformation + vbApplicationModal, "Multiple File Renamer")
        End If
    End If
    GoTo Above
Error_1:
        If Temporaer_Language = "GER" Then
            Message = MsgBox("Datei kann nicht umbenannt werden!", vbCritical + vbApplicationModal, "Multiple File Renamer")
            End
        Else
            Message = MsgBox("File can't be renamed", vbCritical + vbApplicationModal, "Multiple File Renamer")
            End
        End If
        GoTo Above
ENDE:   File1.Refresh
        SelectButton.Enabled = True
        NewName.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Select Case Temporaer_Language
        Case "GER"
            LabelTime.Caption = "Zeit: " & Time
        Case "ENG"
            LabelTime.Caption = "Time: " & Time
    End Select
End Sub

Private Sub Combo1_Click()
    Select Case Combo1.ListIndex
        Case 0 'Alle Dateien
            File1.Pattern = "*.*"
            Count_Listindex = 0
        Case 1 'Textdateien
            File1.Pattern = "*.txt;*.asc;*.rtf;*.ini"
            Count_Listindex = 1
        Case 2 'Bilder
            File1.Pattern = "*.jpg;*.gif;*.png;*.tif;*.bmp;*.swf;*.svg"
            Count_Listindex = 2
        Case 3 'PDFs
            File1.Pattern = "*.pdf"
            Count_Listindex = 3
        Case 4 'Excel-Dateien
            File1.Pattern = "*.xls;*.xlsx"
            Count_Listindex = 4
        Case 5 'Word-Dateien
            File1.Pattern = "*.doc;*.docx"
            Count_Listindex = 5
        Case 6 'Powerpoint-Dateien
            File1.Pattern = "*.ppt;*.pptx"
            Count_Listindex = 6
        Case 7 'Videos
            File1.Pattern = "*.wmv;*.mpg;*.mp4;*.avi;*.mov;*.swf;*.rm;*.vob;*.mkv;*.mpg"
            Count_Listindex = 7
        Case 8 'Musik
            File1.Pattern = "*.mid;*.mp3;*.ogg;*.wav"
            Count_Listindex = 8
    End Select
End Sub

Public Sub picShowPicture(oPictureBox As Object, _
  ByVal sFile As String, _
  Optional ByVal bStretch As Boolean = True)
 
  With oPictureBox
    If bStretch Then
      ' Bild an Größe der PictureBox anpassen
      .AutoRedraw = True
      Set .Picture = Nothing
      .PaintPicture LoadPicture(sFile), 0, 0, .ScaleWidth, .ScaleHeight
      .AutoRedraw = False
    Else
      ' PictureBox an Bildgröße anpassen
      Set .Picture = Nothing
      .Picture = LoadPicture(sFile)
      .AutoSize = True
    End If
  End With
End Sub

Private Sub cmdExtractIcons(Counting_Pics As Integer, Picture_Box As PictureBox)
     Dim nIcon As Long, lRet As Long, lIconCount As Long
     Dim i As Integer, iPicCount As Integer
     Dim sPathToIconFile As String
     Dim hIconArray() As Long
     Dim Icon As Long
   
     sPathToIconFile = "%SystemRoot%\system32\SHELL32.dll"
   
     ' Prüfen, ob die Datei überhaupt Icons enthält.
     ' Wenn ja, Anzahl ermitteln:
     nIcon = -1
     lIconCount = ExtractIcon(App.hInstance, sPathToIconFile, nIcon)
     If lIconCount = 0 Then
       MsgBox "Die ausgewählte Datei enthält keine Icons"
       Exit Sub
     End If
     Icon = ExtractIcon(App.hInstance, sPathToIconFile, Counting_Pics)
     Picture_Box.Cls
     lRet = DrawIcon(Picture_Box.hdc, 0, 0, Icon)
     DoEvents
   End Sub

Public Sub URLGoTo(ByVal hWnd As Long, ByVal URL As String)
  ' hWnd: Das Fensterhandle des aufrufenden Formulars
  Screen.MousePointer = 11
  Call ShellExecute(hWnd, "Open", URL, "", "", 1)
  Screen.MousePointer = 0
End Sub
