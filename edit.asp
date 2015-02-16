 <%
ID=Request.Form("ID")
if ID="" then response.end
    ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open ConnString

Recordset.Open "Select * from datanew where ID=" & ID , Connection
%>
<!DOCTYPE html>
<html>
<head>
<title>ADO - Edit DataBase Record</title>
</head>
<body>
<h2>Edit Database Table</h2>
<form method="post" action="sd.asp" target="_blank">
<input name="ID" type="hidden" value=<%=ID%>>
<table bgcolor="#b0c4de">
<%
for each x in Recordset.Fields
      if x.name <> "ID" and x.name <> "dateadded" then%>
           <tr>
           <td><%=x.name%> </td>
           <td><input name="<%=x.name%>" value="<%=x.value%>" size="20"></td>
      <%end if
next

Recordset.close
Connection.close
%>
</tr>
</table>
<br>
<input type="submit" name="action" value="Save">
<input type="submit" name="action" value="Delete">
</form>

<p><a href="showaspcode.asp?source=demo_db_edit">View source on how to create input fields based on the fields from one record in the database table</a>.</p>
<p><b>Note:</b> If you click on "Save" or "Delete" a new page will open. On the new page you may look at the source on how to submit changes to, or delete from a database table.</p>
</body>
</html>
