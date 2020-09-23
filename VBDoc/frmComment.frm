VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBDoc Subroutine Commenter"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   410
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   476
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frmComment.frx":0000
      Top             =   3285
      Width           =   6135
   End
   Begin VB.ComboBox cboDate 
      Height          =   315
      ItemData        =   "frmComment.frx":0004
      Left            =   900
      List            =   "frmComment.frx":0006
      TabIndex        =   26
      Top             =   2565
      Width           =   2085
   End
   Begin VB.TextBox txtTitle 
      Height          =   330
      Left            =   900
      MultiLine       =   -1  'True
      TabIndex        =   23
      Text            =   "frmComment.frx":0008
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Comment Borders"
      Height          =   2040
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   6990
      Begin VB.OptionButton optHeader 
         Caption         =   "«««««««««««««««"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "»»»»»»»»»»»»»»»"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   540
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "###############"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   810
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "---------------"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   19
         Top             =   1080
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "==============="
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   1350
         Value           =   -1  'True
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "***************"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   1620
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "~~~~~~~~~~~~~~~"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   4545
         TabIndex        =   16
         Top             =   1350
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "_______________"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   4545
         TabIndex        =   15
         Top             =   1080
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "\\\\\\\\\\\\\\\\"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   2340
         TabIndex        =   14
         Top             =   540
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "////////////////"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   2340
         TabIndex        =   13
         Top             =   810
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "###############"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   2340
         TabIndex        =   12
         Top             =   270
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "::::::::::::::::"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   2340
         TabIndex        =   11
         Top             =   1080
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "<<<<<<<<<<<<<<<<"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   2340
         TabIndex        =   10
         Top             =   1350
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   ">>>>>>>>>>>>>>>>"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   2340
         TabIndex        =   9
         Top             =   1620
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "^^^^^^^^^^^^^^^^"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   4545
         TabIndex        =   8
         Top             =   540
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "................"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   4545
         TabIndex        =   7
         Top             =   810
         Width           =   1995
      End
      Begin VB.OptionButton optHeader 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   16
         Left            =   4545
         TabIndex        =   6
         Top             =   1620
         Width           =   240
      End
      Begin VB.TextBox txtHeader 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4860
         TabIndex        =   5
         Top             =   1620
         Width           =   1680
      End
      Begin VB.OptionButton optHeader 
         Caption         =   "ººººººººººººººº"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   17
         Left            =   4545
         TabIndex        =   4
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Proceed"
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   5760
      Width           =   1275
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   5760
      TabIndex        =   1
      Top             =   5760
      Width           =   1275
   End
   Begin VB.TextBox txtComments 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   3960
      Width           =   6945
   End
   Begin VB.Label Label1 
      Caption         =   "Suffix:"
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   29
      Top             =   3330
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   27
      Top             =   2565
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Subroutine Comment Preview:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   25
      Top             =   3735
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "Prefix:"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   24
      Top             =   2205
      Width           =   735
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim sHeaderChar$

Private Sub cboDate_Click()
    Update_Sample sHeaderChar
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub

Private Sub Setup_Dates()
    cboDate.Clear
    cboDate.AddItem "No Date"
    cboDate.AddItem "mmmm dd, yyyy hh:mm:ss"
    cboDate.AddItem "mmmm dd, yyyy hh:mm"
    cboDate.AddItem "mmmm dd, yyyy"
    cboDate.AddItem "mmm dd, yyyy hh:mm:ss"
    cboDate.AddItem "mmm dd, yyyy hh:mm"
    cboDate.AddItem "mmm dd, yyyy"
    cboDate.AddItem "mm/dd/yyyy hh:mm:ss"
    cboDate.AddItem "mm/dd/yyyy hh:mm"
    cboDate.AddItem "mm/dd/yyyy"
    cboDate.AddItem "mm/dd/yy hh:mm:ss"
    cboDate.AddItem "mm/dd/yy hh:mm"
    cboDate.AddItem "mm/dd/yy"
    cboDate.ListIndex = 1
    cboDate_Click
End Sub

Private Sub Form_Load()
    Setup_Dates
    optHeader_Click 4
End Sub

Private Sub optHeader_Click(Index As Integer)
    sHeaderChar = Mid$(optHeader(Index).Caption, 1, 1)
    Update_Sample sHeaderChar
End Sub

Public Sub Update_Sample(ByVal s$)
    Dim t$
    
    txtComments.Text = ""
    Add_Comment Chars(s$, 60)
    Add_Comment "Public Sub Procedure_Name(parameters)"
    Add_Comment ""
    Add_Comment txtTitle.Text
    
    If cboDate.ListIndex > 0 Then
        Add_Comment Format$(Now(), cboDate.List(cboDate.ListIndex))
    End If
    
    Add_Comment ""
    Add_Comment Chars(s$, 60)
    
End Sub

Public Sub Add_Comment(Optional s$ = " ")
    Dim n&, c$, t$
    t$ = ""
    For n = 1 To Len(s$)
        c$ = Mid$(s$, n, 1)
        If c$ = vbCrLf Or c$ = Chr$(13) Or c$ = Chr$(10) Then
            If Len(t$) > 0 Then Print_Comment t$
            t$ = ""
        Else
            t$ = t$ & c$
        End If
    Next n
    If Len(t$) > 0 Then Print_Comment t$
End Sub

Private Sub Print_Comment(Optional comment$ = "")
    txtComments.Text = txtComments.Text & "'" & comment$ & vbCrLf
End Sub

Public Function Chars(ByVal dupeChar$, ByVal length&) As String
    Dim n&, s$
    s$ = ""
    For n = 1 To length
        s$ = s$ & dupeChar$
    Next n
    Chars = s$
End Function

Private Sub txtHeader_KeyPress(KeyAscii As Integer)
    Dim n&, s$
    s$ = ""
    For n = 1 To 20
        s$ = s$ & Chr$(KeyAscii)
    Next n
    optHeader(16).Caption = s$
    txtHeader.Text = s$
    optHeader_Click 16
End Sub

Private Sub txtTitle_Change()
    Update_Sample sHeaderChar
End Sub
