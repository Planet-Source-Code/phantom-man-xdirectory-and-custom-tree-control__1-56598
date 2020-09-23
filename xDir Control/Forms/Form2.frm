VERSION 5.00
Object = "*\A..\xDirectory\xDirectory.vbp"
Begin VB.Form Form2 
   Caption         =   "Main Menu"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7590
   LinkTopic       =   "Form2"
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackGroundPicture=   "Form2.frx":0000
      BorderStyle     =   1
      BackColor       =   16777215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      Height          =   3015
      Left            =   3480
      TabIndex        =   3
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Selection:"
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   960
      Width           =   3375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Main Menu"
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
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   5175
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   6360
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "xDirectory Main Menu"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   135
      Width           =   5175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.xDirectory1.Init , edt_Custom
    
    With Me.xDirectory1
        .CustomAdd 0, "Main Menu", True, "Main Menu", True
        .RefreshCustomNodes
        
    End With
    
        
        
    
    
End Sub


Private Sub xDirectory1_BeforeCustomNodeExpand(lNode As Long, Caption As String, State As Boolean)
Dim i As Long
    
    Dim sDescriptionOne As String
    Dim sDescriptionTwo As String
    
    sDescriptionOne = vbCrLf & "This Sample Shows You The Basics Of A Folder Tree Structure."
    sDescriptionTwo = vbCrLf & "This Sample Shows You How Easy It Is To Make The Tree Do What You Want it To Do." & _
                vbCrLf & vbCrLf & "Here I Am Making A Font Tree Sample."
    
    If lNode = 1 Then
        Me.xDirectory1.CustomAdd lNode, "Directory Tree Sample", False, sDescriptionOne, False
        Me.xDirectory1.CustomAdd lNode, "Font Tree Sample", False, sDescriptionTwo, False
    End If
    
    Me.xDirectory1.RefreshCustomNodes
End Sub

Private Sub xDirectory1_NodeDblClick(ByVal ID As Long, ByVal Caption As String, ByVal sTag As String)
    
   
    Select Case Caption
        Case "Directory Tree Sample"
            Screen.MousePointer = vbHourglass
            Form1.Show vbModal, Me
        Case "Font Tree Sample"
            Screen.MousePointer = vbHourglass
            Form3.Show vbModal, Me
    End Select
    
End Sub

Private Sub xDirectory1_RightClick(ByVal sValue As String)
    MsgBox "Right Clicked On: " & sValue
End Sub

Private Sub xDirectory1_SelectionChanged(ByVal sValue As String, ByVal sTag As String)
    Me.Label2.Caption = "Selection: " & sValue
    Me.Label1.Caption = sTag
End Sub

