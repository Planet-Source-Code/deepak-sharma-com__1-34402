VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COM"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   3360
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "COM With Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1680
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "COM Without Database"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ASP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Visual Basic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label war 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   2520
      Width           =   75
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   14  'Copy Pen
      X1              =   120
      X2              =   5400
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   14  'Copy Pen
      X1              =   120
      X2              =   5400
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   14  'Copy Pen
      X1              =   3000
      X2              =   3000
      Y1              =   720
      Y2              =   2280
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      DrawMode        =   14  'Copy Pen
      X1              =   120
      X2              =   5400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderStyle     =   3  'Dot
      BorderWidth     =   2
      DrawMode        =   14  'Copy Pen
      FillColor       =   &H00C0E0FF&
      Height          =   3735
      Left            =   120
      Top             =   120
      Width           =   5295
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COM Examples"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   360
      Left            =   1440
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim str


Private Sub Command1_Click()
If str = "without" Then
  Shell App.Path & "\Normal COM\Samples\COM Test.exe", vbNormalFocus
  End
ElseIf str = "with" Then
  Shell App.Path & "\Database COM\Samples\COM Test.exe", vbNormalFocus
  End
End If
End Sub

Private Sub Command2_Click()
If str = "without" Then
   
   MsgBox "Copy " & "COM Test.asp" & " from " & App.Path & "\Normal COM\Samples to your virtual directory.", vbInformation
  
ElseIf str = "with" Then
  
  MsgBox "Copy " & "database_COM.asp" & " from " & App.Path & "\Database COM\Samples to your virtual directory.", vbInformation
  
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Option1_Click()
str = "without"
war.Caption = ""
End Sub

Private Sub Option2_Click()
str = "with"
war.Caption = "Make Sure Contact.mdb database exist in c: drive. It is in" & vbCrLf & vbCrLf _
               & App.Path & "\Database COM\Samples\Database"
End Sub
