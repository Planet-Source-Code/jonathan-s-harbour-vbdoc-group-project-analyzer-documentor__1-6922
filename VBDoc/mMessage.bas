Attribute VB_Name = "mMessage"

Option Explicit

'*********************************************************************
'Critical
'Displays a message with a critical icon to indicate a serious warning
'*********************************************************************
Public Sub Critical(ByVal title$, ByVal msg$)
    MsgBox msg$, vbOKOnly + vbCritical, title$
End Sub

'*********************************************************************
'Warning
'Displays a message with an exclamation icon to indicate a warning
'*********************************************************************
Public Sub Warning(ByVal title$, ByVal msg$)
    MsgBox msg$, vbOKOnly + vbExclamation, title$
End Sub

'*********************************************************************
'Message
'Displays a message with an information icon
'*********************************************************************
Public Sub Message(ByVal title$, ByVal msg$)
    MsgBox msg$, vbOKOnly + vbInformation, title$
End Sub


