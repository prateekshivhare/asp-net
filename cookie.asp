<%@ Language="VBScript"%>

<%
    dim novisit
    response.Cookies("visits").Expires=date+365
    novisit= request.cookies("visits")
        if novisit="" then
            response.Cookies("visits")=1
            response.write("Welcome! This is the first time you are visiting this Web page. ")   
        else
            response.Cookies("visits")=novisit+1
            response.Write("you have visited this ")
            response.write(" webpage" & novisit)
        if novisit = 1 then
           response.write(" time before")
          
        else
            response.write(" times before")
    end if
   end if

  %>
<!DOCTYPE html>
<html>
<body>
</body>
</html>




