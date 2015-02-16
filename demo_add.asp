<html>
<body>

<%
  Dim Connection
Dim ConnString
Dim Recordset
Dim SQL
ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"

Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open ConnString
Recordset.Open SQL,Connection



sql="INSERT INTO data (name,lastname,"
sql=sql & "fathername,mothername)"
sql=sql & " VALUES "
sql=sql & "('" & Request.Form("custid") & "',"
sql=sql & "'" & Request.Form("compname") & "',"
sql=sql & "'" & Request.Form("contname") & "',"
sql=sql & "'" & Request.Form("address") & "',"
sql=sql & "'" & Request.Form("city") & "',"
sql=sql & "'" & Request.Form("postcode") & "',"
sql=sql & "'" & Request.Form("country") & "')"
    Connection.Execute(sql)

Recordset.Close
Connection.close
%>

</body>
</html> 