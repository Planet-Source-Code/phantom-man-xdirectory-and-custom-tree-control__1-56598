VERSION 5.00
Object = "*\A..\xDirectory\xDirectory.vbp"
Begin VB.Form Form3 
   Caption         =   "Font Viewer"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form3"
   ScaleHeight     =   4290
   ScaleWidth      =   7590
   StartUpPosition =   2  'CenterScreen
   Begin xDirectoryCTL.xDirectory xDirectory1 
      Align           =   3  'Align Left
      Height          =   4290
      Left            =   0
      TabIndex        =   0
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
      BackGroundPicture=   "Form3.frx":0000
      BorderStyle     =   1
      BackColor       =   16777215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   3600
      TabIndex        =   6
      Top             =   2040
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   3600
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection:"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   3600
      TabIndex        =   2
      Top             =   2880
      Width           =   5415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Font Viewer"
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
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Line Line1 
      X1              =   3720
      X2              =   6600
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Font Viewer"
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
      TabIndex        =   5
      Top             =   135
      Width           =   5175
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.xDirectory1.Init , edt_Fonts
    Screen.MousePointer = vbDefault
End Sub

Private Sub xDirectory1_RightClick(ByVal sValue As String)
    MsgBox "Right Clicked On: " & sValue
End Sub

Private Sub xDirectory1_SelectionChanged(ByVal sValue As String, ByVal sTag As String)
On Error Resume Next
   
   Me.Label1.Caption = sTag
   Me.Label5.Font.Italic = False
   Me.Label5.Font.Bold = False
   
   If sValue <> "Available Fonts" Then
        Me.Label5.FontName = sValue
        Me.Label5.Caption = "This Is A Sample Text"
    Else
        Me.Label5.FontName = Me.FontName
        Me.Label5.Caption = ""
    End If
    
On Error GoTo 0
    
    
End Sub

