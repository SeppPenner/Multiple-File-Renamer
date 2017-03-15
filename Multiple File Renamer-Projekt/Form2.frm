VERSION 5.00
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8175
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8175
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LicenseBox 
      Height          =   6690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    LicenseBox.Clear
    LicensePath = App.Path & "\License.txt"
    Set License = FSO.OpenTextFile(LicensePath, ForReading)
    While Not License.AtEndOfStream
        DoEvents
        LicenseBox.AddItem License.ReadLine
        DoEvents
        LicenseBox.Refresh
        DoEvents
    Wend
    If Temporaer_Language = "GER" Then
        Form2.Caption = "Multiple File Renamer-Lizenz"
    Else
        Form2.Caption = "Multiple File Renamer-License"
    End If
End Sub
