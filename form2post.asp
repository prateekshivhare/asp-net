<%@ Language="VBScript"%>

<%
Dim mySessionID
mySessionID = Session.SessionID
response.write("the session id generated is" & Session.SessionID)
%>
<%
dim novisit
response.Cookies("visits").Expires=date+365
novisit= request.cookies("visits")
if novisit="" then
response.Cookies("visits")=1
response.write("Welcome! This is the first time you are visiting this Web page.")
else
response.Cookies("visits")=novisit+1
response.Write("you have visited this ")
response.write(" webpage " & novisit)
if novisit = 1 then
response.write(" time before")
else
response.write(" times before")
end if
end if
%>
<!DOCTYPE html>
<html>
<body bgcolor="#E0F5FF" font-size: 40px;>
<form action="update.asp" method="post">
<table align="center">
<tr>
<td>
Name:</td> <td> <input type="text" name="fname" size="20" /></td></tr>
<tr><td>Lastname:</td> <td> <input type ="text" name="lname" size="20"/></td></tr>
<tr><td>Father's Name:</td> <td> <input type ="text" name="fathrname" size="20"/></td></tr>
<tr><td>Mother's Name:</td> <td> <input type ="text" name="mothrname" size="20"/></td></tr>
<tr><td>Sex: <input type="radio" name="sex"
<%
if sex="male" then Response.Write("checked") %>
value="male">male</input>
<input type="radio" name="sex"
<%if sex="female" then Response.Write("checked")%>
value="female">female</input></td></tr>
<tr><td><input type="submit" value="Submit" /></td></tr>
</table>
</form>
</body>
</html>