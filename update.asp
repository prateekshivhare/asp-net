 <%option explicit%>
<!DOCTYPE html>
<html>
<head><title>ADO - List Database Records</title></head>
<body>

    <%
Dim Connection
Dim ConnString
Dim Recordset
Dim SQL
ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open ConnString


    sql="INSERT INTO  datanew (name,lastname,fathername,mothername) "
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("fname") & "',"
sql=sql & "'" & Request.Form("lname") & "',"
sql=sql & "'" & Request.Form("fathrname") & "',"
sql=sql & "'" & Request.Form("mothrname") & "')"

Connection.Execute SQL




Connection.Close

%>



<%



Dim x
ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"
SQL = "SELECT * FROM datanew"
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open ConnString
 Recordset.Open SQL,Connection

%>
<h2>List Database Table</h2>
<p>Click on a button to modify a record.</p>

<table border="1" width="100%">
<tr bgcolor="#b0c4de">
<%
for each x in Recordset.Fields
      response.write("<th>" & ucase(x.name) & "</th>")
next
%>
</tr>
<%do until Recordset.EOF%>
<tr bgcolor="#f0f0f0">
<form method="post" action="edit.asp" target="_blank">
<%
for each x in Recordset.Fields
      if x.name="ID" then%>
           <td><input type="submit" name="ID" value="<%=x.value%>"></td>
      <%else%>
           <td><%Response.Write(x.value)%> </td>
      <%end if
next
%>
</form>
<%Recordset.MoveNext%>
</tr>
<%
loop
Recordset.close
set Recordset=nothing
Connection.close
set Connection=nothing
%>
</table>

<p><a href="showaspcode.asp?source=demo_db_list">View source on how to list a database table in an HTML table</a></p>
<p><b>Note:</b> If you click on a button in the "no" column a new page will open. On the new page you may look at the source on how to create input fields based on the fields from one record in the database table.</p>
</body>
</html>
