VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Functions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-------------------------------------
'COM Example   : For Beginers
'Developed By  : Deepak Sharma
'E-Mail        : deepakmailto@rediffmail.com
'-------------------------------------------

'------------------------------------------
'              NOTE
'             ------
'
' Please Vote me and give me your feedbacks
' so that i can judge my knowledge
'
'------------------------------------------


Dim param1, param2, Functions

Public Function ADD(first As String, second As String)
  
  param1 = first
  param2 = second
  Functions = "ADD"
  
  If ErrorHandler(Functions) = "" Then ADD = Val(first) + Val(second)
    
End Function

Public Function MINUS(first As String, second As String)
  
  param1 = first
  param2 = second
  Functions = "MINUS"
  
  If ErrorHandler(Functions) = "" Then MINUS = Val(first) - Val(second)
  
End Function

Public Function MUL(first As String, second As String)
  
  param1 = first
  param2 = second
  Functions = "MUL"
  
  If ErrorHandler(Functions) = "" Then MUL = Val(first) * Val(second)
    
End Function

Public Function DIV(first As String, second As String)
  
  param1 = first
  param2 = second
  Functions = "DIV"
  
 If ErrorHandler(Functions) = "" Then DIV = Val(first) / Val(second)
  
End Function

Private Function ErrorHandler(ByVal Function_Name As String)
    
   If Function_Name = "ADD" Or Function_Name = "MUL" Or _
      Function_Name = "MINUS" Or Function_Name = "DIV" Then
       
       If param1 = "" Then
         ErrorHandler = "Enter First Number"
         Exit Function
       End If
               
       If param2 = "" Then
         ErrorHandler = "Enter Second Number"
         Exit Function
       End If
        
   End If
  
  
   If Function_Name = "DIV" Then
  
      If param2 <> "" Then
        If param2 = "0" Then
           ErrorHandler = "Cannot Division By Zero"
           Exit Function
        End If
      End If
  
   End If
  
  
   If Function_Name = "MINUS" Then
    
     If Val(param1) < Val(param2) Then
        ErrorHandler = "First Number Cannot Be Less Then The Second Number"
        Exit Function
     End If
  
   End If

End Function

Public Function ERRORS()

  ERRORS = ErrorHandler(Functions)
  
End Function

