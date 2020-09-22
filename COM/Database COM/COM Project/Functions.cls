VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------
'COM Example   : For Advance
'Developed By  : Deepak Sharma
'E-Mail        : deepakmailto@rediffmail.com
'-------------------------------------------

'------------------------------------------
'              NOTE
'             ------
'
' Please Vote me and give me your feedbacks
' so that i can judge my knowledge
'------------------------------------------

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim Max As New ADODB.Recordset
Dim Find As New ADODB.Recordset
Dim incre

Private Sub Class_Initialize()

   If cn.State = 1 Then cn.Close
   'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database\contact.mdb;Persist Security Info=False"
   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\contact.mdb;Persist Security Info=False"
   cn.CursorLocation = adUseClient
   rs.Open "select * from customer", cn, adOpenDynamic, adLockOptimistic
 
End Sub

Public Property Let Code(NewValue As Integer)

    rs("CODE") = NewValue

End Property

Public Property Let Names(NewValue As String)

    rs("Name") = NewValue

End Property

Public Property Let Age(NewValue As Integer)

    rs("Age") = NewValue

End Property

Public Property Let Phone(NewValue As Long)

    rs("Phone") = NewValue

End Property

Public Property Let Married(NewValue As String)

    rs("Married") = IIf(NewValue = "Yes", True, False)

End Property

Public Property Get Code() As Integer

    Code = rs("CODE")

End Property

Public Property Get Names() As String

    Names = rs("Name")

End Property

Public Property Get Age() As Integer

    Age = rs("Age")

End Property

Public Property Get Phone() As Long

    Phone = rs("Phone")

End Property

Public Property Get Married() As String
  
  Married = IIf(rs.Fields("married") = True, "Yes", "No")
  
End Property

Public Property Get EOF() As Boolean

  EOF = rs.EOF
  
End Property

Public Property Get BOF() As Boolean

  BOF = rs.EOF
  
End Property

Public Function MoveNext()
 
    rs.MoveNext

End Function

Public Function MoveLast()
 
    rs.MoveLast
  
End Function

Public Sub MovePrev()
  
  rs.MovePrevious
  
End Sub

Public Sub MoveFirst()

  rs.MoveFirst
  
End Sub

Public Sub AddNewRecord()
 
  rs.AddNew
 
End Sub

Public Sub SaveRecord()
 
  rs.Update
 
End Sub

Public Function DeleteRecord(Cust_Code As Integer)
 
  cn.Execute "delete from customer where code=" & Cust_Code
 
End Function

Public Function Update(Cust_Code As Integer)
 
  If rs.State = 1 Then rs.Close
  rs.Open "select * from customer where code=" & Cust_Code, cn, adOpenDynamic, adLockOptimistic
 
End Function

Public Function GetCustID() As Integer
 
   If Max.State = 1 Then Max.Close
   Max.Open "select max(code) from customer", cn, adOpenDynamic, adLockOptimistic
   incre = Max.Fields(0)
   
   If IsNull(incre) Then
     GetCustID = 1
   ElseIf Not IsNull(incre) Then
     GetCustID = Max.Fields(0) + 1
   End If
  
End Function

Public Function FindCustomer(Optional Cust_Code As Integer, Optional Cust_Name As String) As Boolean

   If Cust_Code > 0 Then
    
      If rs.State = 1 Then rs.Close
      rs.Open "select * from customer where code=" & Cust_Code, cn, adOpenDynamic, adLockOptimistic
      
      If rs.RecordCount < 1 Then
         FindCustomer = False
      Else
         FindCustomer = True
      End If
    
   ElseIf Not IsMissing(Cust_Name) Then
      
      If rs.State = 1 Then rs.Close
      rs.Open "select * from customer where name=" & "'" & Cust_Name & "'", cn, adOpenDynamic, adLockOptimistic
      If rs.RecordCount < 1 Then
         FindCustomer = False
      ElseIf rs.RecordCount >= 1 Then
         FindCustomer = True
      End If
      
   End If
  
End Function

Public Sub ResetSQL()

  If rs.State = 1 Then rs.Close
  rs.Open "select * from customer", cn, adOpenDynamic, adLockOptimistic

End Sub

Public Function CurrentPosition() As Long

   CurrentPosition = rs.AbsolutePosition
   
End Function

Public Function TotalRecords() As Long
  
   TotalRecords = rs.RecordCount
   
End Function

Private Sub Class_Terminate()

    cn.Close
    Set cn = Nothing

End Sub
