VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HCString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'----------------------------------------------------------------------------------------
' Copyright © 1999 by Jonathan Harbour.  All rights reserved.
'
' Project:      Harbour Class Library
' Class:        HCString
' Description:  String manipulation routines
'
' Date          Author      Comments
' 12/02/99      JSH         Initial code completed
'
'----------------------------------------------------------------------------------------

Option Explicit

Private Enum OpMode
  StringBinaryCompare = vbBinaryCompare + 1
  StringTextCompare = vbTextCompare + 1
  StringDataBaseCompare = vbDatabaseCompare + 1
  CharacterBinaryCompare = -(vbBinaryCompare + 1)
  CharacterTextCompare = -(vbTextCompare + 1)
  CharacterDataBaseCompare = -(vbDatabaseCompare + 1)
End Enum

Public Function IsUpper(ByVal letter$) As Boolean
    IsUpper = True
    If Asc(letter$) < 65 Or Asc(letter$) > 90 Then
        IsUpper = False
    End If
End Function

Public Function IsLower(ByVal letter$) As Boolean
    IsLower = True
    If Asc(letter$) < 97 Or Asc(letter$) > 122 Then
        IsLower = False
    End If
End Function

Public Function ParseStr(ByVal strWork$, ByVal intTokenNum&, _
    ByVal strDelimitChr$, Optional ByVal strEncapChr$) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Retrieve specified token of string.
  '
  ' ARGUMENTS
  '     strWork       = String to work on.
  '     intTokenNum   = If > 0,  returns specified token in string. If 0, returns
  '                     next token in string each time function is called. (If
  '                     no more tokens are found, function will return 0.) To
  '                     reset counter to 0, call routine as ParseStr ("", 0, "").
  '
  '     strDelimitChr = Token delimiter.
  '     strEncapChr   = Optional. Allows for tokens to return strings
  '                     encapsulated with "strDelimitChr" characters.
  '
  ' RETURNS
  '     Returns string token.  If none is found, will return "".
  '
  ' DEPENDENCIES
  '     ReplaceCS
  '
  ' REMARKS
  '
  '     If you are in "auto-mode" (intTokenNum=0) and are going to auto
  '     process another string, make sure to reset it as follows:
  '
  '     CALL ParseStr("", 0, "") or ParseStr "", 0, ""
  '
  '*******************************************************************************
    
  On Local Error Resume Next
  
  Dim blnExitDo As Boolean
  Dim intDPos As Integer
  Dim intSPtr As Integer
  Dim intEPtr As Integer
  Dim intCurrentTokenNum As Integer
  Dim intWorkStrLen As Integer
  Dim intEncapStatus As Integer
  Static intSPos As Integer
  Dim strTemp As String
  Static intDelimitLen As Integer

  intWorkStrLen = Len(strWork)
    
  If Len(strEncapChr) Then
    intEncapStatus = Len(strEncapChr)
  End If

  If intWorkStrLen = 0 Or (intSPos > intWorkStrLen And intTokenNum = 0) Then
    intSPos = 0
    Exit Function
  ElseIf intTokenNum > 0 Or intSPos = 0 Then
    intSPos = 1
    intDelimitLen = Len(strDelimitChr)
  End If

  Do
    
    intDPos = InStr(intSPos, strWork, strDelimitChr)

    If intEncapStatus Then
      intSPtr = InStr(intSPos, strWork, strEncapChr)
      intEPtr = InStr(intSPtr + 1, strWork, strEncapChr)
      If intDPos > intSPtr And intDPos < intEPtr Then
        intDPos = InStr(intEPtr, strWork, strDelimitChr)
      End If
    End If

    If intDPos < intSPos Then
      intDPos = intWorkStrLen + intDelimitLen
    End If

    If intDPos Then
      If intTokenNum Then
        intCurrentTokenNum = intCurrentTokenNum + 1
        If intCurrentTokenNum = intTokenNum Then
          strTemp = Mid(strWork, intSPos, intDPos - intSPos)
          blnExitDo = True
        Else
          blnExitDo = False
        End If
      Else
        strTemp = Mid(strWork, intSPos, intDPos - intSPos)
          blnExitDo = True
      End If
      intSPos = intDPos + intDelimitLen
    Else
      intSPos = 0
      blnExitDo = True
    End If
  Loop Until blnExitDo

  If intEncapStatus Then
    ParseStr = ReplaceCS(strTemp, strEncapChr, "", StringBinaryCompare)
  Else
    ParseStr = strTemp
  End If

End Function


Private Function ReplaceCS(ByVal strWork As String, ByVal strOld As String, _
    ByVal strNew As String, Optional ByVal intOPMode As OpMode, _
    Optional blnUpdated As Boolean) As String

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Replace a string or specific character(s) within a string. This routine
  '     can also be used to strip characters.
  '
  ' ARGUMENTS
  '     strWork    = String to work on.
  '
  '     strOld     = If intOPMode is negative, defines character(s) to replace.
  '                = If intOPMode is positive, defines string to replace.
  '
  '     strNew     = New character (or string) to substitute.
  '
  '     intOPMode  = Sets operation by defining the "replace" mode and "compare"
  '                  mode. Valid parameters are:
  '
  '                  StringBinaryCompare (Default if not specified.)
  '                  StringTextCompare
  '                  StringDataBaseCompare
  '                  CharacterBinaryCompare
  '                  CharacterTextCompare
  '                  CharacterDataBaseCompare
  '
  '     blnUpdated = Optional. Returns TRUE if string was modified
  '
  ' RETURNS
  '     Returns new string.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '     To strip a string or character(s), set strNew to vbNullString or "".
  '
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim intOldLen As Integer
  Dim intNewLen As Integer
  Dim intSPos As Long
  Dim intN As Integer
  
  If intOPMode = 0 Then
    intOPMode = StringBinaryCompare
  End If
  
  intNewLen = Len(strNew)
  intOldLen = Len(strOld)
    
  intSPos = 1
  blnUpdated = False
    
  If intOPMode < 0 Then
    intOPMode = Abs(intOPMode) - 1
    For intN = 1 To intOldLen
      intSPos = 1
      Do
        intSPos = InStr(intSPos, strWork, Mid(strOld, intN, 1), intOPMode)
        If intSPos Then
          strWork = Left(strWork, intSPos - 1) & strNew & Mid(strWork, intSPos + 1)
          intSPos = intSPos + intNewLen
          blnUpdated = True
        End If
      Loop While intSPos
    Next
  Else
    intOPMode = intOPMode - 1
    Do
      intSPos = InStr(intSPos, strWork, strOld, intOPMode)
      If intSPos Then
        strWork = Left(strWork, intSPos - 1) & strNew & Mid(strWork, intSPos + intOldLen)
        intSPos = intSPos + intNewLen
        blnUpdated = True
      End If
    Loop While intSPos
  End If
  
  ReplaceCS = strWork
  
End Function


Public Function TokenCount(strWork As String, strDelimiter As String) As Long

  '*******************************************************************************
  '
  ' DESCRIPTION
  '     Counts number of tokens in a string.  Function can also be used to
  '     determine the number of "delimiter" characters found in the string by
  '     subtracting 1 from the returned value.
  '
  ' ARGUMENTS
  '     strWork      = String to work on
  '     strDelimiter = String Delimiter
  '
  ' RETURNS
  '     Number of tokens found.
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
  Dim lngN As Long
  Dim lngCPos As Long
  Dim lngSPos As Long
  Dim lngCharLen As Long
    
  If Len(strWork) = 0 Then Exit Function
  
  lngCharLen = Len(strDelimiter)
  lngSPos = 1
    
  Do
    lngCPos = InStr(lngSPos, strWork, strDelimiter)
    If lngCPos Then
      lngN = lngN + 1
      lngSPos = lngCharLen + lngCPos
    End If
  Loop While lngCPos
   
  If Right(strWork, lngCharLen) <> strDelimiter Then
    TokenCount = lngN + 1
  Else
    TokenCount = lngN
  End If
      
End Function


Public Function ProperCase(sInput As String) As String
    'code written by Chris Cubley, May 1998
    Dim iCnt As Integer
    Dim boolNewWord As Boolean

    boolNewWord = True

    For iCnt = 1 To Len(sInput)
        If boolNewWord = True Then
            Mid(sInput, iCnt, 1) = UCase(Mid(sInput, iCnt, 1))
        Else
            Mid(sInput, iCnt, 1) = LCase(Mid(sInput, iCnt, 1))
        End If

        If Mid(sInput, iCnt, 1) = " " Then
            boolNewWord = True
        Else
            boolNewWord = False
        End If
    Next iCnt

    ProperCase = sInput
End Function

Public Sub CopyComboList(FromCtl As Variant, ToCtl As Variant, Optional strMode As String)
    
  '*******************************************************************************
  '
  ' DESCRIPTION
  '    Copies or moves the contents of a list/combo box to another
  '
  ' ARGUMENTS
  '    FromCtl - Source List/Combo box control.
  '    ToCtl   - Destination List/Combo box control.
  '    StrMode - Optional. When not specified, copies selected item(s).
  '                  "-" : Same as above but removes selected item(s) after copying.
  '                "ALL" : Copies all items.
  '               "-ALL" : Copies and removes all items after copying.
  '
  ' RETURNS
  '     N/A
  '
  ' DEPENDENCIES
  '     None
  '
  ' REMARKS
  '     Please note that because you can't select multiple items in a combo,
  '     control, all items will processed.
  '
  '*******************************************************************************
  
  On Local Error Resume Next
  
    Dim intN As Integer
    
    If Not (TypeOf FromCtl Is ListBox Or TypeOf ToCtl Is ComboBox) Then
        Debug.Print "CopyComboList: controls must be ListBox or ComboBox!"
        Exit Sub
    End If
    
    With FromCtl
      If TypeName(FromCtl) = "ListBox" Then
        For intN = .ListCount - 1 To 0 Step -1
          If .Selected(intN) Or InStr(strMode, "ALL") Then
            ToCtl.AddItem .List(intN)
            ToCtl.ItemData(ToCtl.NewIndex) = .ItemData(intN)
            If InStr(strMode, "-") = 1 Then .RemoveItem (intN)
          End If
        Next
      Else
        For intN = .ListCount - 1 To 0 Step -1
          ToCtl.AddItem .List(intN)
          ToCtl.ItemData(ToCtl.NewIndex) = .ItemData(intN)
          If InStr(strMode, "-") = 1 Then .RemoveItem (intN)
        Next
      End If
    End With
    
End Sub



