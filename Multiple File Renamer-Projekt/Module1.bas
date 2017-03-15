Attribute VB_Name = "Module1"
Public Language As String
Public LanguagePath As String
Public FSO As New FileSystemObject
Dim Inp As TextStream
Public Temporaer_Language As String
Public Count_Language_Fail As Boolean
Public Count_Listindex As Integer
Public PathWithoutFile As String
Public PathWithFile As String
Public FileEnding As String
Public TargetPath As String
Public Count_Files As Integer
Public Count_Help As Integer
Dim License As TextStream
Public LicensePath As String
Public ALTPressed As Boolean
Public vbKeyAlt
