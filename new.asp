<%@ Language="VBScript"%>
<%
Session.Timeout = 240
Response.Write("The timeout is: " & Session.Timeout)
Response.Write("The session id id: " & Session.SessionID)
%>
<%
dim fname
dim sex
dim lname
dim fathrname
dim mothrname
fname=Request.Form("fname")
sex =Request.Form("sex")
lname =Request.Form("lname")
fathrname =Request.Form("fathrname")
mothrname =Request.Form("mothrname")
If fname<>"" Then
Response.Write("name: " & fname & " " & lname & "<br>")
end If
if sex<>"" then
Response.Write("<p>sex: " & sex & "<br>")
end if
If fathrname<>"" Then
Response.Write("fathername: " & fathrname & "<br>")
End If
If mothrname<>"" Then
Response.Write("Mothername: " & mothrname & "<br>")
End If
%>