VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VBDoc"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   8985
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdComment 
      Caption         =   "&Comment"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1710
      TabIndex        =   65
      Top             =   5850
      Width           =   1320
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4500
      Top             =   5805
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "&Load"
      Height          =   375
      Left            =   90
      TabIndex        =   2
      Top             =   5850
      Width           =   1320
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   5850
      Width           =   1320
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   3735
      Width           =   8790
   End
   Begin VB.Label Label1 
      Caption         =   "Const:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   30
      Left            =   3465
      TabIndex        =   64
      Top             =   2655
      Width           =   1350
   End
   Begin VB.Label lblConst 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   63
      Top             =   2655
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Type:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   29
      Left            =   3465
      TabIndex        =   62
      Top             =   2385
      Width           =   1350
   End
   Begin VB.Label lblTypes 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   61
      Top             =   2385
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   7830
      TabIndex        =   60
      Top             =   3330
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   7800
      TabIndex        =   59
      Top             =   3060
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   7800
      TabIndex        =   58
      Top             =   2790
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   7800
      TabIndex        =   57
      Top             =   2520
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   7830
      TabIndex        =   56
      Top             =   2250
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   7800
      TabIndex        =   55
      Top             =   1980
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   7800
      TabIndex        =   54
      Top             =   1710
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   7800
      TabIndex        =   53
      Top             =   1440
      Width           =   840
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   7830
      TabIndex        =   52
      Top             =   1170
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   7830
      TabIndex        =   51
      Top             =   900
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   7830
      TabIndex        =   50
      Top             =   630
      Width           =   795
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7830
      TabIndex        =   49
      Top             =   360
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Other:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   28
      Left            =   6435
      TabIndex        =   48
      Top             =   3330
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Object:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   27
      Left            =   6435
      TabIndex        =   47
      Top             =   3060
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Variant:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   26
      Left            =   6435
      TabIndex        =   46
      Top             =   2790
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "String:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   25
      Left            =   6435
      TabIndex        =   45
      Top             =   2520
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   24
      Left            =   6435
      TabIndex        =   44
      Top             =   2250
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Decimal:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   23
      Left            =   6435
      TabIndex        =   43
      Top             =   1980
      Width           =   1350
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   7815
      TabIndex        =   42
      Top             =   90
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Byte:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   22
      Left            =   6435
      TabIndex        =   41
      Top             =   90
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Boolean:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   21
      Left            =   6435
      TabIndex        =   40
      Top             =   360
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Integer:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   20
      Left            =   6435
      TabIndex        =   39
      Top             =   630
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Long:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   6435
      TabIndex        =   38
      Top             =   900
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Single:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   18
      Left            =   6435
      TabIndex        =   37
      Top             =   1170
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Double:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   17
      Left            =   6435
      TabIndex        =   36
      Top             =   1440
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Currency:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   16
      Left            =   6435
      TabIndex        =   35
      Top             =   1710
      Width           =   1350
   End
   Begin VB.Label lblDesigners 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   34
      Top             =   2025
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Designers:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   15
      Left            =   3465
      TabIndex        =   33
      Top             =   2025
      Width           =   1350
   End
   Begin VB.Label lblUserDocuments 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   32
      Top             =   1755
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "User Documents:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   14
      Left            =   3465
      TabIndex        =   31
      Top             =   1755
      Width           =   1350
   End
   Begin VB.Label lblPropertyPages 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   30
      Top             =   1485
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Property pages:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   13
      Left            =   3465
      TabIndex        =   29
      Top             =   1485
      Width           =   1350
   End
   Begin VB.Label lblUserControls 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   28
      Top             =   1215
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "User controls:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   12
      Left            =   3465
      TabIndex        =   27
      Top             =   1215
      Width           =   1350
   End
   Begin VB.Label lblClassModules 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   26
      Top             =   945
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Class modules:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   11
      Left            =   3465
      TabIndex        =   25
      Top             =   945
      Width           =   1350
   End
   Begin VB.Label lblCodeModules 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   24
      Top             =   675
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Code modules:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   3465
      TabIndex        =   23
      Top             =   675
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Forms:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   3465
      TabIndex        =   22
      Top             =   405
      Width           =   1350
   End
   Begin VB.Label lblForms 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4845
      TabIndex        =   21
      Top             =   405
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Total subroutines:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   225
      TabIndex        =   20
      Top             =   1155
      Width           =   1455
   End
   Begin VB.Label lblTotalSubroutines 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   19
      Top             =   1155
      Width           =   795
   End
   Begin VB.Label lblCommentLines 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   18
      Top             =   885
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Comment lines:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   465
      TabIndex        =   17
      Top             =   885
      Width           =   1410
   End
   Begin VB.Label lblSourceLines 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   630
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Source code lines:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   450
      TabIndex        =   15
      Top             =   630
      Width           =   1410
   End
   Begin VB.Label Label1 
      Caption         =   "Total lines of code:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   180
      TabIndex        =   14
      Top             =   375
      Width           =   1680
   End
   Begin VB.Label lblTotalLines 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   13
      Top             =   375
      Width           =   795
   End
   Begin VB.Label lblPrivateFunctions 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   12
      Top             =   2100
      Width           =   795
   End
   Begin VB.Label lblPublicFunctions 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   1845
      Width           =   795
   End
   Begin VB.Label lblPrivateSubs 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   1590
      Width           =   795
   End
   Begin VB.Label lblPublicSubs 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   1380
      Width           =   795
   End
   Begin VB.Label lblProjectName 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1935
      TabIndex        =   8
      Top             =   90
      Width           =   4035
   End
   Begin VB.Label Label1 
      Caption         =   "Project file:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   180
      TabIndex        =   7
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Private functions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   465
      TabIndex        =   6
      Top             =   2100
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Public functions:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   465
      TabIndex        =   5
      Top             =   1845
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Private subs:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   465
      TabIndex        =   4
      Top             =   1590
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Public subs:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   465
      TabIndex        =   3
      Top             =   1380
      Width           =   1395
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------
' Program Title:    VBDoc
' Author:           Jonathan S. Harbour
' Date:             11/30/99
'
' Revisions:    DATE        AUTHOR      COMMENTS
'               11/30/99    JSH         Initial program started
'               12/09/99    JSH         HTML output
'
'-------------------------------------------------------------------------------

Option Explicit

Dim html As New HCHTML
Dim strng As New HCString

Dim subs As New Collection
Dim sCurObject$

Dim iTotalPrivSubs&, iTotalPubSubs&, iTotalPrivFuncs&, iTotalPubFuncs&
Dim iTotalConsts&, iTotalLongs&, iTotalInt&, iTotalBool&, iTotalSingle&, iTotalDouble&
Dim iTotalByte&, iTotalCurrency&, iTotalDecimal&, iTotalObject&, iTotalDate&, iTotalString&
Dim iTotalVariant&, iTotalOther&, iTotalTypes&, iTotalCommentLines&

Private Sub cmdComment_Click()
    frmComment.Show vbModal
End Sub

Private Sub cmdLoad_Click()
    CD1.DialogTitle = "Load Project"
    CD1.Filter = "Visual Basic Project|*.vbp;*.vbg"
    CD1.InitDir = App.Path
    CD1.Orientation = cdlLandscape
    CD1.ShowOpen
    If Len(CD1.fileName) > 0 Then
        txtOutput.Text = ""
        Load_Project CD1.fileName
    End If
End Sub

Public Sub Load_Project(ByVal fileName$)
    Dim filenum&, s$, count&
    
    If InStr(1, fileName$, ".") = 0 Then
        Critical App.title, "Invalid filename: " & fileName$
        Exit Sub
    End If
    
    html.Create App.Path & "\vbdoc.html"
    html.Pr "<TT><FONT SIZE=""3"">"
    html.PrLn "<B>VBDoc Generated Listing</B></FONT>"
    html.Pr "<HR ALIGN=""CENTER""><FONT SIZE=""2"">"
    html.PrLn "Version: " & App.Major & "." & App.Minor & ", Build: " & App.Revision
    html.PrLn "Author: Jonathan S. Harbour"
    html.PrLn "Generated: " & Now()
    html.PrLn "<HR ALIGN=""CENTER"">"
    
    Disp "Project File: " & fileName$
    lblProjectName.Caption = fileName$
    
    count = Count_Projects(fileName$)
    If count > 0 Then Disp "Projects: " & FmtNum(count)
    
    count = Count_Objects(fileName$, "Form")
    Disp "Forms: " & FmtNum(count)
    lblForms.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "Module")
    Disp "Code Modules: " & FmtNum(count)
    lblCodeModules.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "Class")
    Disp "Class Modules: " & FmtNum(count)
    lblClassModules.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "UserControl")
    Disp "User Controls: " & FmtNum(count)
    lblUserControls.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "PropertyPage")
    Disp "Property Pages: " & FmtNum(count)
    lblPropertyPages.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "UserDocument")
    Disp "User Documents: " & FmtNum(count)
    lblUserDocuments.Caption = FmtNum(count)
    
    count = Count_Objects(fileName$, "Designer")
    Disp "Designers: " & FmtNum(count)
    lblDesigners.Caption = FmtNum(count)
    
    Disp "Line Count:"
    count = Count_TotalLines(fileName$)
    Disp "Total Lines: " & FmtNum(count)
    Disp "  Comment Lines: " & FmtNum(iTotalCommentLines)
    Disp "  Source Lines: " & FmtNum(count - iTotalCommentLines)
    
    lblTotalLines.Caption = FmtNum(count)
    lblCommentLines.Caption = FmtNum(iTotalCommentLines)
    lblSourceLines.Caption = FmtNum(count - iTotalCommentLines)
    
    Disp "Declares:"
    Disp "  Type: " & FmtNum(iTotalTypes)
    Disp "  Const: " & FmtNum(iTotalConsts)
    
    lblTypes.Caption = FmtNum(iTotalTypes)
    lblConst.Caption = FmtNum(iTotalConsts)
    
    Disp "Variables:"
    Disp "  Byte: " & FmtNum(iTotalByte)
    Disp "  Boolean: " & FmtNum(iTotalBool)
    Disp "  Integer: " & FmtNum(iTotalInt)
    Disp "  Long: " & FmtNum(iTotalLongs)
    Disp "  Single: " & FmtNum(iTotalSingle)
    Disp "  Double: " & FmtNum(iTotalDouble)
    Disp "  Currency: " & FmtNum(iTotalCurrency)
    Disp "  Decimal: " & FmtNum(iTotalDecimal)
    Disp "  Date: " & FmtNum(iTotalDate)
    Disp "  String: " & FmtNum(iTotalString)
    Disp "  Variant: " & FmtNum(iTotalVariant)
    Disp "  Object: " & FmtNum(iTotalObject)
    Disp "  Other: " & FmtNum(iTotalOther)
    
    lblType(0).Caption = FmtNum(iTotalByte)
    lblType(1).Caption = FmtNum(iTotalBool)
    lblType(2).Caption = FmtNum(iTotalInt)
    lblType(3).Caption = FmtNum(iTotalLongs)
    lblType(4).Caption = FmtNum(iTotalSingle)
    lblType(5).Caption = FmtNum(iTotalDouble)
    lblType(6).Caption = FmtNum(iTotalCurrency)
    lblType(7).Caption = FmtNum(iTotalDecimal)
    lblType(8).Caption = FmtNum(iTotalDate)
    lblType(9).Caption = FmtNum(iTotalString)
    lblType(10).Caption = FmtNum(iTotalVariant)
    lblType(11).Caption = FmtNum(iTotalObject)
    lblType(12).Caption = FmtNum(iTotalOther)
    
    Disp "Subroutines:"
    Disp "  Private Subs: " & FmtNum(iTotalPrivSubs)
    Disp "  Public Subs: " & FmtNum(iTotalPubSubs)
    Disp "  Private Functions: " & FmtNum(iTotalPrivFuncs)
    Disp "  Public Functions: " & FmtNum(iTotalPubFuncs)
    Disp "  Total Subroutines: " & FmtNum(iTotalPrivSubs + iTotalPubSubs + iTotalPrivFuncs + iTotalPubFuncs)
    Disp ""
    
    lblPublicSubs.Caption = FmtNum(iTotalPubSubs)
    lblPrivateSubs.Caption = FmtNum(iTotalPrivSubs)
    lblPublicFunctions.Caption = FmtNum(iTotalPubFuncs)
    lblPrivateFunctions.Caption = FmtNum(iTotalPrivFuncs)
    lblTotalSubroutines.Caption = FmtNum(iTotalPrivSubs + iTotalPubSubs + iTotalPrivFuncs + iTotalPubFuncs)
    
    'list project file
    filenum = FreeFile()
    Open fileName$ For Input As #filenum
    Do Until EOF(filenum)
        Line Input #filenum, s$
'        Disp s$
    Loop
    Close #filenum

    html.PrLn ""
    html.PrLn "<HR ALIGN=""CENTER"">"
    html.PrLn "</FONT></TT><FONT SIZE=""2"">Copyright 1999 <A HREF=""mailto:jsharbour@home.com"">" & "Jonathan S. Harbour" & "</A></FONT>"
    html.PrLn "Visit the <A HREF=""http://24.5.57.182/vbdoc"">VBDoc Web Site</A> for more information."
    html.CloseFile
    
End Sub

Public Function FmtNum(ByVal num&)
    FmtNum = Format$(num, "#,##0")
End Function

Public Function Count_Projects(ByVal fileName$) As Long
    Dim filenum&, total&, s$
    
    total = 0
    filenum = FreeFile()
    Open fileName$ For Input As #filenum
    Do Until EOF(filenum)
        Line Input #filenum, s$
        If InStr(1, s$, "Project=") Or InStr(1, s$, "StartupProject=") Then
            total = total + 1
        End If
    Loop
    Close #filenum
    Count_Projects = total
End Function

Public Function Count_Objects(ByVal fileName$, ByVal obj$) As Long
    Dim filenum&, total&, s$
    
    total = 0
    filenum = FreeFile()
    Open fileName$ For Input As #filenum
    Do Until EOF(filenum)
        Line Input #filenum, s$
        If InStr(1, s$, "Project=") Or InStr(1, s$, "StartupProject=") Then
            total = total + Count_Objects(strng.ParseStr(s$, 2, "="), obj$)
        ElseIf InStr(1, s$, "Startup" & obj$ & "=") = 1 Then
            'ignore this case
        ElseIf InStr(1, s$, "Icon" & obj$ & "=") = 1 Then
            'ignore this case
        ElseIf InStr(1, s$, obj$ & "=") = 1 Then
            total = total + 1
        End If
    Loop
    Close #filenum
    Count_Objects = total
End Function

Public Function Count_TotalLines(ByVal fileName$) As Long
    Dim filenum&, total&, s$, fn$, count&
    
    total = 0
    filenum = FreeFile()
    Open fileName$ For Input As #filenum
    Do Until EOF(filenum)
        Line Input #filenum, s$
        If Len(s$) > 0 Then
            
            'count lines
            If InStr(1, s$, "Project=") Or InStr(1, s$, "StartupProject=") Then
                total = total + Count_TotalLines(Trim$(strng.ParseStr(s$, 2, "=")))
            ElseIf InStr(1, s$, "Form=") Then
                fn$ = Trim$(strng.ParseStr(s$, 2, "="))
                If Right$(fn$, 4) = ".frm" Then
                    Disp "  " & fn$
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  Total lines: " & count
                End If
            ElseIf InStr(1, s$, "Module=") Then
                fn$ = Trim$(strng.ParseStr(strng.ParseStr(s$, 2, "="), 2, ";"))
                If Right$(fn$, 4) = ".bas" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            ElseIf InStr(1, s$, "Class=") Then
                fn$ = Trim$(strng.ParseStr(strng.ParseStr(s$, 2, "="), 2, ";"))
                If Right$(fn$, 4) = ".cls" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            ElseIf InStr(1, s$, "UserControl=") Then
                fn$ = Trim$(strng.ParseStr(s$, 2, "="))
                If Right$(fn$, 4) = ".ctl" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            ElseIf InStr(1, s$, "PropertyPage=") Then
                fn$ = Trim$(strng.ParseStr(s$, 2, "="))
                If Right$(fn$, 4) = ".pag" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            ElseIf InStr(1, s$, "UserDocument=") Then
                fn$ = Trim$(strng.ParseStr(s$, 2, "="))
                If Right$(fn$, 4) = ".dob" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            ElseIf InStr(1, s$, "Designer=") Then
                fn$ = Trim$(strng.ParseStr(s$, 2, "="))
                If Right$(fn$, 4) = ".dsr" Then
                    count = Count_Lines(fn$)
                    total = total + count
                    Disp "  " & fn$ & ": " & count & " lines"
                End If
            End If
        End If
    Loop
    Close #filenum
    Count_TotalLines = total
End Function

Public Function Count_Lines(ByVal fileName$) As Long
    Dim filenum&, total&, s$, n&, token$
    Dim param As Boolean
    
    On Error GoTo error1
    total = 0
    filenum = FreeFile()
    Open fileName$ For Input As #filenum
    Do Until EOF(filenum)
        Line Input #filenum, s$
        
        If Len(s$) > 0 Then
            param = False
            If InStr(1, s$, "Private Sub ") = 1 Then
                token$ = Mid$(s$, Len("Private Sub ") + 1)
                'Disp "      Private Sub: " & sCurObject & "." & token$
                iTotalPrivSubs = iTotalPrivSubs + 1
                param = True
            ElseIf InStr(1, s$, "Public Sub ") = 1 Then
                iTotalPubSubs = iTotalPubSubs + 1
                param = True
            ElseIf InStr(1, s$, "Private Function ") = 1 Then
                iTotalPrivFuncs = iTotalPrivFuncs + 1
                param = True
            ElseIf InStr(1, s$, "Public Function ") = 1 Then
                iTotalPubFuncs = iTotalPubFuncs + 1
                param = True
            ElseIf InStr(1, s$, "Const ") = 1 Then
                iTotalConsts = iTotalConsts + 1
            
            'variable declaration or subroutine parameter
            ElseIf InStr(1, s$, "Dim ") Or param Then
                For n = 1 To strng.TokenCount(s$, ",")
                    token$ = strng.ParseStr(s$, (n), ",")
                    If InStr(1, token$, " As Long") Or InStr(1, token$, "&") Then
                        iTotalLongs = iTotalLongs + 1
                    ElseIf InStr(1, token$, " As Integer") Or InStr(1, token$, "%") Then
                        iTotalInt = iTotalInt + 1
                    ElseIf InStr(1, token$, " As Boolean") Then
                        iTotalBool = iTotalBool + 1
                    ElseIf InStr(1, token$, " As Byte") Then
                        iTotalByte = iTotalByte + 1
                    ElseIf InStr(1, token$, " As Single") Or InStr(1, token$, "!") Then
                        iTotalSingle = iTotalSingle + 1
                    ElseIf InStr(1, token$, " As Double") Or InStr(1, token$, "#") Then
                        iTotalDouble = iTotalDouble + 1
                    ElseIf InStr(1, token$, " As Currency") Or InStr(1, token$, "@") Then
                        iTotalCurrency = iTotalCurrency + 1
                    ElseIf InStr(1, token$, " As Decimal") Then
                        iTotalDecimal = iTotalDecimal + 1
                    ElseIf InStr(1, token$, " As Date") Then
                        iTotalDate = iTotalDate + 1
                    ElseIf InStr(1, token$, " As String") Or InStr(1, token$, "$") Then
                        iTotalString = iTotalString + 1
                    ElseIf InStr(1, token$, " As Object") Then
                        iTotalObject = iTotalObject + 1
                    ElseIf InStr(1, token$, " As Variant") Then
                        iTotalVariant = iTotalVariant + 1
                    ElseIf InStr(1, token$, " As ") And Len(token$) > 4 Then
                        iTotalOther = iTotalOther + 1
                    End If
                Next n
            
            'user defined types
            ElseIf InStr(1, s$, "Type ") = 1 Then
                iTotalTypes = iTotalTypes + 1
            
            'VB_Name holds name of object (form, module, etc)
            ElseIf InStr(1, s$, "Attribute VB_Name = ") = 1 Then
                sCurObject = Trim$(strng.ParseStr(s$, 2, "="))
                sCurObject = Mid$(sCurObject, 2, Len(sCurObject) - 2)   'strip quotes
                
            'comment lines
            ElseIf Mid$(Trim$(s$), 1, 1) = "'" Then
                iTotalCommentLines = iTotalCommentLines + 1
            
            End If
        End If
        total = total + 1
    Loop
    Close #filenum
    Count_Lines = total
    Exit Function
error1:
    Close #filenum
    Critical "Count_Lines", Err.Description
End Function

Private Sub cmdQuit_Click()
    End
End Sub

Public Sub Disp(ByVal s$)
    txtOutput.Text = txtOutput.Text & s$ & vbCrLf
    html.PrLn s$
End Sub

Private Sub Form_Load()
    Me.Show
    cmdLoad_Click
End Sub



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


