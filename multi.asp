<%@ Language="VBScript" %>		
<html>
<head>
<title>My First ASP Page</title>
</head>
<body bgcolor="white" text="black">
 	<%
         sub multi(n1,n2)
         response.write(n1*n2)
         end sub         
         
          %>
 		<p>
             result :<%call multi(3,4)  %></p>
	</body>
</html>
	