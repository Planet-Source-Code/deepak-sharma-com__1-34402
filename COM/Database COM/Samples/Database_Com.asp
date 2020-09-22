<%
'**********************************************************
' Design          : DEEPAK SHARMA
' Developed       : DEEPAK SHARMA
' Level           : Advance
' COM Name        : Database.Function
'
' THIS SAME COM COMPONENT ( Database.Function ) IS RUNNING
' IN VISUAL BASIC UNDER DATABASE COM FOLDER 
'**********************************************************
%>
<%@language=vbscript%>
<%option explicit%>
<%Response.Buffer=true%>
<%
Dim code,name,married,age,phone,func,caption,str,Flag,ctl
 
 'CREATE THE COM OBJECT
 
 set func=server.createobject("Database.function")        
 
 
 if caption="" then caption="ADD"   
 
  call FindCust 'search through combobox
 
  select case Request.Form("operation")    
  
  Case "CLEAR"
  
		  call Clear_Fields
      
  Case "ADD"                  
         
          call Clear_Fields
          code = Func.GetCustID
          caption="SAVE"                 
         
  Case "SAVE"
        
            Func.ResetSQL  'reset the Sql query
                    
            Func.AddNewRecord            
            
            Func.Code = cint(Request.Form("code"))
            Func.Names = Request.Form("name")
            Func.Age = cint(Request.Form("age"))
            Func.Married = Request.Form("married")
            Func.Phone = CLng (Request.Form("phone"))
            
            Func.SaveRecord
            
            Func.ResetSQL              
            
            call Message("Record Is Saved")           		
  
    case "MODIFY"
    
            Func.Update cint(Request.Form("code")) 'accecpt the customer code to modify
    
            Func.Names = Request.Form("name")
            Func.Age = cint(Request.Form("age"))
            Func.Married = Request.Form("married")
            Func.Phone = CLng (Request.Form("phone"))           
    
            Func.SaveRecord
            
            move_fields
            
            Func.ResetSQL            
            
            call Message("Record Is Update")            
            
  
    case "DELETE"
           
           Func.DeleteRecord Request.Form("code") 'accecpt the customer code to delete
    
           call message("Record is Deleted")
           
           call clear_fields
           
           Func.ResetSQL       
  
    case "JUMP"
         
           If Request("Adv_search")<>"" then
            
             If Func.FindCustomer(, Trim(Request("Adv_search"))) = True Then
                 
                 Move_Fields
                 func.ResetSQL                  
                 
             Else                 
                 
                 call message("No Record Found")                 
                 func.ResetSQL 
                 
             End If     
             
           end if
           
   end select
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><BR><br>

<body bgcolor="#ffffff" >
<form name=myform method=post action="database_com.asp">
<table width="50%" bgcolor=ivory border="0" cellpadding=2 cellspacing=2 id=TABLE1 ALIGN=CENTER >
  <tr bgcolor=brown> 
    <td colspan="2" height="31"> 
      <div align="center"><font color=yellow face="Arial, Helvetica, sans-serif" size=4><b>COM Database 
        Example</b></font></div>
    </td>
  </tr>
  <tr>
    <td width="42%" height="24"><font size="3" face=arial><b> Search</b></font></td>
    <td width="58%" height="24"> 
      <select name="search" style="HEIGHT: 22px; WIDTH: 226px" onchange="myform.submit()">            
      <%While Not Func.EOF = True%>
  	   <option <%if cint(Request.Form("search"))=cint(func.Code) then Response.Write "selected"%> value=<%=func.Code%>><%=Func.Names%></option>
	   <%Func.Movenext%>
       <%Wend%>
       <%Func.movefirst%>            
      </select>
    </td>
  </tr>
  <tr> 
    <td width="42%"><font size="3" face=arial><b>Code</b></font></td>
    <td width="58%"> 
      <input name="code" style="HEIGHT: 22px; WIDTH: 79px" value=<%=Code%>>
    </td>
  </tr>
  <tr> 
    <td width="42%"><font size="3" face=arial><b>Name</b></font></td>
    <td width="58%"> 
      <input type=text name="name" style="HEIGHT: 22px; WIDTH: 224px"  value=<%=replace(server.URLEncode(name),"+","")%>>
    </td>
  </tr>
  <tr> 
    <td width="42%"><font size="3" face=arial><b>Married</b></font></td>
    <td width="58%"> 
      <select name="married" style="HEIGHT: 22px; WIDTH: 66px" >
      <option <%if married="Yes" then Response.Write "selected"%>>Yes</option>
      <option <%if married="No" then Response.Write "selected"%>>No</option>      
      </select>
    </td>
  </tr>
  <tr> 
    <td width="42%"><font size="3" face=arial><b>Age</b></font></td>
    <td width="58%"> 
      <input name="age" style="HEIGHT: 22px; WIDTH: 71px"  value=<%=age%>>
    </td>
  </tr>
  <tr> 
    <td width="42%"><font size="3" face=arial><b>Phone</b></font></td>
    <td width="58%"> 
      <input name="phone" style="HEIGHT: 22px; WIDTH: 71px"  value=<%=phone%>>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;             
    </td>
  </tr>
  
  <tr> 
    <td width="42%"><font size="3" face=arial color=blue><b>Enter Name</b></font></td>
    <td width="58%"> 
      <input name="Adv_search" size=30>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;       
    </td>
  </tr>
  
  <tr bgcolor=brown> 
    <td colspan="2">    
      <P> <br>
      <center>
      <input type="submit"  name="operation"  value=<%=caption%> style="HEIGHT: 24px; WIDTH: 61px">
      <input type="submit"  name="operation"  value=MODIFY >
      <INPUT type="submit"  name="operation"  value=DELETE>      
      <INPUT type="submit"  name="operation"  value=CLEAR>
      <INPUT type="submit"  name="operation"  value=JUMP>
      </center>
      </P>      
    </td>
  </tr>
</table>
</form>
</body>
</html>
<%
public sub Clear_Fields()

  Code=""
  name=""
  married=""
  age=""
  phone=""
  
End sub


public sub Move_fields()  
      
    Code=func.Code     
	name=func.Names 
    married=func.Married 
    age=func.Age 
    phone=func.Phone    
    
End sub      
 
  
public sub FindCust()
  
    if Request.Form("Search")<>"" then
      
	  if func.FindCustomer(Request.Form("Search")) then	
	
         call Move_fields
    
         with Response
           .Write "<script language=javascript>"          
           .Write "myform.submit()"
           .Write "</script>"    
         end with  

      end if      
      func.ResetSQL             
    
    end if 
    
End sub
 
 
 public sub Message(str)
 
   with Response
    .Write "<script language=javascript>"
    .Write "alert("+"'" &  str &"'"+");"
    .Write "myform.submit();"
    .Write "</script>"    
   end with  
   
 End sub   

%>