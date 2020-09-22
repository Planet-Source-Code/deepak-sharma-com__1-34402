VERSION 5.00
Begin VB.Form frmtest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "COM Example"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   4080
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcode 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton CMDCANCEL 
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      TabIndex        =   17
      Top             =   2760
      Width           =   1335
      Begin VB.CommandButton cmdnext 
         Caption         =   "Next"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "Prev"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.CommandButton CMDCLOSE 
      Caption         =   "CLOSE"
      Height          =   255
      Left            =   2040
      TabIndex        =   16
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "CLEAR"
      Height          =   255
      Left            =   1080
      TabIndex        =   15
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton cmdjump 
      Caption         =   "JUMP"
      Height          =   255
      Left            =   3000
      TabIndex        =   14
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      Height          =   255
      Left            =   2040
      TabIndex        =   13
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton cmdmod 
      Caption         =   "MODIFY"
      Height          =   255
      Left            =   1080
      TabIndex        =   12
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox txtphone 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Tag             =   "Enter Phone"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.TextBox txtage 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Tag             =   "Enter Age"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cbostatus 
      Height          =   315
      ItemData        =   "frmTest.frx":0000
      Left            =   1080
      List            =   "frmTest.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Tag             =   "Select Married Status"
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtname 
      Height          =   285
      Left            =   1080
      TabIndex        =   8
      Tag             =   "Enter Name"
      Top             =   1800
      Width           =   2775
   End
   Begin VB.ComboBox cbosearch 
      Height          =   315
      Left            =   1080
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Code"
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
      TabIndex        =   22
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Married"
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
      Top             =   2400
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Phone"
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
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Age"
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
      TabIndex        =   3
      Top             =   2880
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Name"
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
      TabIndex        =   2
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "COM Database Example"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3465
   End
End
Attribute VB_Name = "frmtest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Func As Database.Function
Dim str$

Private Sub cbosearch_Click()
    
  On Error Resume Next
    
  If Func.FindCustomer(Me.cbosearch.ItemData(Me.cbosearch.ListIndex)) Then
     Move_Fields
  End If
    
  Func.ResetSQL

End Sub

Private Sub CMDCANCEL_Click()

    Me.cmdadd.Caption = "ADD"
    cmdclear_Click
    Me.cmdnext.Visible = False
    Me.cmdprev.Visible = False
    Me.CMDCANCEL.Enabled = False
    
End Sub

Private Sub CMDCLOSE_Click()
 
  End
  
End Sub

Private Sub cmddelete_Click()
        
    Func.DeleteRecord Val(Me.txtcode.Text)
    
    MsgBox "Record is Deleted", vbInformation
    cmdclear_Click
    Func.ResetSQL
    FillSearchBox
    Me.cmdnext.Visible = False
    Me.cmdprev.Visible = False
    
End Sub

Private Sub cmdjump_Click()
  
  str = InputBox(vbCrLf + "Enter Customer Name : ", "Advance Search")
  If str <> "" Then
    If Func.FindCustomer(, Trim(str)) = True Then
       Me.cmdnext.Visible = True
       cmdnext_Click
       Move_Fields
    Else
       MsgBox "No Record Found", vbInformation
    End If
  End If
 
End Sub

Private Sub cmdmod_Click()
    
    Func.Update Val(txtcode.Text)
    
    Func.Names = Me.txtname.Text
    Func.Age = Val(Me.txtage.Text)
    Func.Married = Me.cbostatus.Text
    Func.Phone = Val(Me.txtphone.Text)
    
    Func.SaveRecord
    
    MsgBox "Record is Updated"
    Func.ResetSQL
    FillSearchBox
    Me.cmdnext.Visible = False
    Me.cmdprev.Visible = False
    
End Sub

Private Sub cmdnext_Click()
    
    If Func.CurrentPosition = Func.TotalRecords Then
 
        Me.cmdnext.Visible = False
           
    Else
     
        Me.cmdprev.Visible = True
        Func.Movenext
        Move_Fields
        
    End If
      
End Sub

Private Sub cmdadd_Click()
     
   If Me.cmdadd.Caption = "ADD" Then
      
        cmdclear_Click
        Me.txtcode.Text = Func.GetCustID
        Me.cmdadd.Caption = "SAVE"
        Me.CMDCANCEL.Enabled = True

    ElseIf Me.cmdadd.Caption = "SAVE" Then

        If Blank_Fields = False Then
     
            Func.ResetSQL
            
            Func.AddNewRecord
            
            Func.Code = Val(Me.txtcode.Text)
            Func.Names = Me.txtname.Text
            Func.Age = Val(Me.txtage.Text)
            Func.Married = Me.cbostatus.Text
            Func.Phone = Val(Me.txtphone.Text)
            
            Func.SaveRecord
            
            MsgBox "One row(s) added to the database", vbInformation
            
            Me.cmdadd.Caption = "ADD"
            
            Func.ResetSQL
            FillSearchBox
            
        End If
        
    End If

End Sub

Private Sub cmdclear_Click()

  Me.cmdadd.Caption = "ADD"
  With frmtest
    .txtcode.Text = ""
    .txtname.Text = ""
    .txtage.Text = ""
    .txtphone.Text = ""
    .cbosearch.Text = " "
  End With

End Sub

Private Sub cmdprev_Click()

   If Func.CurrentPosition = 2 Then

      Me.cmdprev.Visible = False
  
   End If
  
   Func.MovePrev
   Move_Fields
   Me.cmdnext.Visible = True

End Sub

Private Sub Form_Load()

  Set Func = New Database.Function
  
  FillSearchBox
  
End Sub


Public Sub Move_Fields()

  With frmtest
    .txtcode.Text = Func.Code
    .txtname.Text = Func.Names
    .cbostatus.Text = Func.Married
    .txtage.Text = Func.Age
    .txtphone.Text = Func.Phone
  End With
  
End Sub

Public Sub FillSearchBox()
  
  Me.cbosearch.Clear
  
  While Not Func.EOF = True
   Me.cbosearch.AddItem Func.Names
   Me.cbosearch.ItemData(Me.cbosearch.NewIndex) = Func.Code
   Func.Movenext
  Wend
  
  'This will add the blank line after filling all record
  
  Me.cbosearch.AddItem " ", Me.cbosearch.ListCount
  
  
  '----------------------------------------------------------------------------
  '                       DO NOT REMOVE THIS LINE
  '----------------------------------------------------------------------------
  'THIS LINE WILL MOVE THE RECORDSET POINTER TO THE FIRST RECORD BECAUSE AFTER
  'FILLING THE COMBO BOX RECORDSET POINTER REACH AT THE LAST RECORD NOW IF YOU
  'PRESS NEXT BUTTON IT WILL GENERATE ERROR BECASUE ITS ALREADY IN LAST RECORD.
  'SO THIS LINE IS USED TO COME ON FIRST RECORD
  '-----------------------------------------------------------------------------
  
  Func.movefirst
  
End Sub

Private Sub txtname_LostFocus()
txtname.Text = StrConv(txtname.Text, vbProperCase)
End Sub

Private Function Blank_Fields() As Boolean

  For Each ctl In Me.Controls
   If TypeOf ctl Is TextBox Then
     If ctl.Tag <> "" Then
       If ctl.Text = "" Then
          MsgBox ctl.Tag, vbCritical
          Blank_Fields = True
          ctl.SetFocus
          Exit For
        Else
          Blank_Fields = False
       End If
     End If
   End If
  Next
  
End Function
