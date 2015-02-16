 <!DOCTYPE html>
<html>
<head>
<title>ADO - Submit DataBase Record</title>
</head>
<body>
<h2>Submit to Database</h2>
     
<%on error resume next
     ConnString="DRIVER={SQL Server};SERVER=jeweldb2012\retail;UID=sa;" & _
"PWD=diaspark@1234;DATABASE=Employees_development"
    Set Connection = Server.CreateObject("ADODB.Connection")
    SET objCommand = Server.CreateObject("ADODB.command")
    

'Set Recordset = Server.CreateObject("ADODB.Recordset")

Connection.Open ConnString

    
    

if Request.form("action")="Save" then
 
    
'      ID=Request.Form("ID")
 '     sql="UPDATE datanew SET name='" & Request.Form("name") & "',"
  '    sql=sql & "lastname='" & Request.Form("lastname") & "',"
   '   sql=sql & "fathername='" & Request.Form("fathrname") & "',"
    '  sql=sql & "mothername='" & Request.Form("mothrname") & "'"
     ' sql=sql & "where ID=" & ID
    

'Connection.Execute sql
    objCommand.ActiveConnection = Connection 
    objCommand.CommandText = "spUpdateEmployee"
    objCommand.CommandType = 4 'adCmdStoredProc
    
    objCommand.Parameters("@UName") =Request.Form("name")
    objCommand.Parameters("@LastName") = Request.Form("lastname")
    objCommand.Parameters("@FName") = Request.Form("fathername")
    objCommand.Parameters("@MName") = Request.Form("mothername")
    objCommand.Parameters("@ID") = Request.Form("ID")
   
    objCommand.Execute
 if err <> 0 then
 Response.Write("Error: " & err.Description & ", Source: " & err.Source)
 'Response.Write("You do not have permission to update this database!")
 else
 Response.Write("Record number " & ID & " was updated.")
 end if
      Response.Write("Submitting records has been disabled from this demo")
end if
if Request.Form("action")="Delete" then
      ID=Request.Form("ID")
Connection.Execute "DELETE FROM datanew WHERE ID=" & ID, Recordsaffected
 if err <> 0 then
 Response.Write("You do not have permission to delete a record from this database!")
 else
 Response.Write("Record number " & ID & " was deleted.")
 end if
      Response.Write("Deleting records has been disabled from this demo")
end if
objCommand.ActiveConnection = nothing

Connection.close%>

<p><a href="showaspcode.asp?source=demo_db_submit">View source on how to save changes to, or delete from a database</a>.</p>
</body>
</html>
