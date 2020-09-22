<%@ Language=VBScript %>
<%
 Dim func,final_result,first_number,second_number

 '----------------------------------
 'create the instance of COM Object
 '---------------------------------- 

 set func=createObject("Operation.Functions")
  
 
 select case Request.Form("oper")
   
   case "ADD"
     
      final_result= func.ADD(Request.Form("first"),Request.Form("second"))     
   
   case "MINUS"
     
      final_result= func.MINUS(Request.Form("first"),Request.Form("second"))
   
   case "MULTIPLY"
     
     final_result= func.MUL(Request.Form("first"),Request.Form("second"))
   
   case "DIVIDE"
     
     final_result= func.DIV(Request.Form("first"),Request.Form("second"))
  
 end select

%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<style type=text/css>  
  .Settings1 {Text-Align:center}
  .Settings2 {Text-Align:center;BACKGROUND-COLOR:wheat;COLOR:RED}
  .Settings3 {Background-color:black;color:#00FF66;font-size:13}  
   Body{margin-top:12%}
</style>
<BODY >
<form method=post name=myform action="com test.asp">
    <table width="39%" border="1" bgcolor=wheat align=center>
    <tr> 
      <td colspan="2"> 
        <div align="center"><b>Com Example</b></div>
      </td>
    </tr>
    <tr> 
      <td width="57%"><font size="2">Enter First Number</font></td>
      <td width="43%"><input type="text" class="settings1" name="first" value="<%=first_number%>"></td>
    </tr>
    <tr> 
       <td width="57%"><font size="2">Enter Second Number</font></td>
       <td width="43%"><input type="text" class="settings1" name="second" value="<%=second_number%>"></td>
    </tr>
    <tr> 
       <td width="57%"><font size="2">Final Result</font></td>
       <td width="43%"><input type="text" READONLY class="settings2" name="result" value="<%=final_result%>"></td>
    </tr>
    <tr > 
      <td width="57%" COLSPAN=2 class="settings3"><center><%=func.Errors%></center></td>
      
    </tr>  
    <tr> 
      <td colspan="2"><center><input type="submit" name="oper" value="ADD">
                      <input type="submit" name="oper" value="MINUS">
                      <input type="submit" name="oper" value="MULTIPLY">
                      <input type="submit" name="oper" value="DIVIDE">            
                      </center></td>
    </tr>
    </table>
</form>
</BODY>
</HTML>
