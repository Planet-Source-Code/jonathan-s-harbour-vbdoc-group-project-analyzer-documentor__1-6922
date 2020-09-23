VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HCHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' Copyright © 1999 by Jonathan Harbour.  All rights reserved.
'
' Project:      Harbour Class Library
' Class:        HCHTML
' Description:  HTML output
'
' Date          Author      Comments
' 12/09/99      JSH         Initial code completed
'
'----------------------------------------------------------------------------------------

Option Explicit

Private sFileName$
Private iFileNum&

Private Sub Class_Initialize()
    sFileName$ = ""
    iFileNum = 0
End Sub

Public Sub Create(ByVal fileName$)
    On Error Resume Next
    sFileName$ = fileName$
    iFileNum = FreeFile
    Open sFileName$ For Output As #iFileNum

    Print #iFileNum, "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 3.2//EN"" > "
    Print #iFileNum, "<!-- Generated on " & Now() & " -->"
    Print #iFileNum, "<HTML>"
    Print #iFileNum, ""
    Print #iFileNum, "<HEAD>"
    Print #iFileNum, "    <META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html;CHARSET=iso-8859-1"">"
    Print #iFileNum, "    <META NAME=""GENERATOR"" Content=""VBDoc"">"
    Print #iFileNum, "    <META NAME=""Author"" Content=""Jonathan S. Harbour"">"
    Print #iFileNum, "<TITLE>VBDoc Generated Listing</TITLE>"
    Print #iFileNum, "</HEAD>"
    Print #iFileNum, ""
    Print #iFileNum, "<BODY TEXT=""#000080"" BGCOLOR=""#FFFFFF"" LINK=""#0000FF"" VLINK=""#0099FF"" ALINK=""#0000FF"">"
    
    Close iFileNum

End Sub

Public Sub Append(ByVal fileName$)
    On Error Resume Next
    sFileName$ = fileName$
    iFileNum = FreeFile
    Open sFileName$ For Append As #iFileNum
    Close iFileNum

End Sub

Public Sub CloseFile()
    PrLn ""
    PrLn "</BODY>"
    PrLn ""
    PrLn "</HTML>"
    PrLn ""
End Sub

Public Sub Output(ByVal line$)
    Dim n&, c$
    
    
    On Error Resume Next
    If Len(sFileName$) = 0 Then Exit Sub
    iFileNum = FreeFile
    Open sFileName$ For Append As #iFileNum
    
    Do While Left$(line$, 1) = Chr$(32)
        Print #iFileNum, "&nbsp;";
        line$ = Mid$(line$, 2)
    Loop
    Print #iFileNum, line$
    
    Close iFileNum
End Sub

Private Sub OutputLn(ByVal s$)
    Output s$ & "<BR>"
End Sub

Public Sub Pr(ByVal s$)
    Output s$
End Sub

Public Sub PrLn(Optional s$)
    If IsMissing(s) Then
        OutputLn ""
    Else
        OutputLn s$
    End If
End Sub


