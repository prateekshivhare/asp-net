<%@ Language="VBScript"%>



<%
Dim Connection
Dim ConnString
Dim Recordset
Dim SQL
    
ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"
SQL = "SELECT * FROM data"
Set Connection = Server.CreateObject("ADODB.Connection")
Set Recordset = Server.CreateObject("ADODB.Recordset")
Connection.Open ConnString
Recordset.Open SQL,Connection
If Recordset.EOF Then
Response.Write("No records returned.")
Else
Do While NOT Recordset.Eof
Response.write Recordset("name")             
Response.write "<br>"
Response.write Recordset("lastname")
Response.write "<br>"
Response.write Recordset("fathername")
Response.write "<br>"
Response.write Recordset("mothername")
Response.write "<br>"
Response.write "<br>"
Recordset.MoveNext
Loop
End If
Recordset.Close
Set Recordset=nothing
Connection.Close
Set Connection=nothing
%>