VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Com"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   -1  'True
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3120
      Width           =   4215
   End
   Begin VB.TextBox txtresult 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CMDRESET 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CMDDIV 
      Caption         =   "DIVIDE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton CMDMULTI 
      Caption         =   "MULTIPLY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CMDSUB 
      Caption         =   "MINUS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CMDADD 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtsecond 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox txtfirst 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Final Result"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Width           =   1020
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "COM Example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   120
      Width           =   1605
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter Second Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   1875
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter First Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   1590
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Func
Dim ks

Private Sub CMDADD_Click()

  'Addition COM function
   Me.txtresult.Text = Func.Add(Me.txtfirst.Text, Me.txtsecond.Text)
  'Handle errors
   txtmsg.Text = Func.errors
   
End Sub

Private Sub CMDCLOSE_Click()
 End
End Sub

Private Sub CMDDIV_Click()

 'Divison COM function
  Me.txtresult.Text = Func.DIV(Me.txtfirst.Text, Me.txtsecond.Text)
 'Handle errors
  txtmsg.Text = Func.errors
  Me.txtfirst.SetFocus
  
End Sub

Private Sub CMDMULTI_Click()

   'Multiply COM function
   Me.txtresult.Text = Func.MUL(Me.txtfirst.Text, Me.txtsecond.Text)
   'Handle errors
   txtmsg.Text = Func.errors
   Me.txtfirst.SetFocus

End Sub

Private Sub CMDRESET_Click()

    Me.txtfirst.Text = ""
    Me.txtresult.Text = ""
    Me.txtsecond.Text = ""
    Me.txtmsg.Text = ""
    Me.txtfirst.SetFocus
    
End Sub

Private Sub CMDSUB_Click()

   'Minus COM function
   Me.txtresult.Text = Func.MINUS(Me.txtfirst.Text, Me.txtsecond.Text)
   'Handle errors
   txtmsg.Text = Func.errors
   Me.txtfirst.SetFocus
 
End Sub

Private Sub Form_Load()

  Set Func = CreateObject("Operation.Functions")
 
End Sub

Private Sub txtfirst_KeyPress(KeyAscii As Integer)

  'For accecpting the integer value
  IntegerValue KeyAscii
 
End Sub

Public Sub IntegerValue(key As Integer)

  ks = key
  If (ks >= 48 And ks <= 57) Or (ks = 32 Or ks = 8) Then
    key = ks
  Else
    key = 0
  End If
  
End Sub

Private Sub txtsecond_KeyPress(KeyAscii As Integer)

  'For accecpting the integer value
  IntegerValue KeyAscii
  
End Sub
