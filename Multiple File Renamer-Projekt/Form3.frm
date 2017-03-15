VERSION 5.00
Begin VB.Form Form3 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Multiple File Renamer"
   ClientHeight    =   1650
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5190
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5190
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox DeveloperBox 
      Height          =   1620
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    DeveloperBox.Clear
    If Temporaer_Language = "GER" Then
        Form3.Caption = "Multiple File Renamer-Entwickler"
        DeveloperBox.AddItem "Produkt: " + CStr(App.ProductName)
        DeveloperBox.AddItem "Produkt Version: " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
        DeveloperBox.AddItem "Beschreibung: Ein Programm, um schnell viele Dateien umzubenennen"
        DeveloperBox.AddItem "Firma: " + CStr(App.CompanyName)
        DeveloperBox.AddItem "Copyright: " + CStr(App.LegalCopyright)
        DeveloperBox.AddItem "Trademarks: " + CStr(App.LegalTrademarks)
        DeveloperBox.AddItem "Entwickler: " + "Tim Hammer"
        DeveloperBox.AddItem "Shortcuts: " + "Alt+U:Umbenennen; Alt+S:Sprache ändern"
    Else
        Form3.Caption = "Multiple File Renamer-Developer"
        DeveloperBox.AddItem "Product: " + CStr(App.ProductName)
        DeveloperBox.AddItem "Product version: " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
        DeveloperBox.AddItem "File description: " + CStr(App.FileDescription)
        DeveloperBox.AddItem "Company: " + CStr(App.CompanyName)
        DeveloperBox.AddItem "Copyright: " + CStr(App.LegalCopyright)
        DeveloperBox.AddItem "Trademarks: " + CStr(App.LegalTrademarks)
        DeveloperBox.AddItem "Developer: " + "Tim Hammer"
        DeveloperBox.AddItem "Shortcuts: " + "Alt+R:Rename; Alt+L:Change Language"
    End If
End Sub

