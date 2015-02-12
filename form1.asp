<%@ Language="VBScript"%>
<html>
<body>
<form method="get"action="form1.asp">
Yourname: <input type ="text" name="fname" size ="20"/>
<input type ="submit" value="submit"/>
</form>
<%
dim fname
fname = Request.QueryString("fname")
if fname<>"" Then
		Response.Write("hello " &fname&"!<br>")
		Response.Write("how are you today?")
end if

%>
</body>
</html>