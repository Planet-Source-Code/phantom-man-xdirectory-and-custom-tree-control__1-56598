VERSION 5.00
Object = "*\A..\xDirectory\xDirectory.vbp"
Begin VB.Form Form1 
   Caption         =   "xDirectory Demo"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin xDirectoryCTL.xDirectory xDirectory1 
      Align           =   3  'Align Left
      Height          =   4290
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   7567
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackGroundPicture=   "Form1.frx":0000
      BorderStyle     =   1
      BackColor       =   16777215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "V1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   6960
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   7440
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Custom Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection:"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   1815
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   3615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Custom Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Index           =   1
      Left            =   3735
      TabIndex        =   4
      Top             =   135
      Width           =   5175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.xDirectory1.Init App.Path, edt_Drives
    Screen.MousePointer = vbDefault
End Sub

Private Sub xDirectory1_LocatingComplete()
    Me.Label3.Caption = "Locating Complete.."
End Sub

Private Sub xDirectory1_Locating(ByVal sDirectory As String)
    Me.Label3.Caption = "Gathering Directory Structure For: " & sDirectory

End Sub

Private Sub xDirectory1_RightClick(ByVal sValue As String)
    MsgBox "Right Clicked On: " & sValue
End Sub

Private Sub xDirectory1_SelectionChanged(ByVal sValue As String, ByVal sTag As String)
    Me.Label1.Caption = sTag
End Sub
