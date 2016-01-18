<%Set Conn = Server.CreateObject("ADODB.Connection")      
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
Set Rs = Server.CreateObject("ADODB.Recordset")
if Request.Form("BtnLogin")="Login"then
   set Max = conn.Execute("select * from idtable where user = '" & Request.Form("Username") & "' and pass = '" & Request.Form("Password") & "'" )
   If max.eof = false then 
     session("Admin") = "Admin" 
     Response.Redirect"admin.asp?Ref=Log" 
   else 
     Response.Redirect"admin.asp?Ref=Try"
   end if 'max=false  
else   
  Ref=Request.Querystring("Ref")
  if Ref="" or Ref="Try" then%>
    <html>
    <head><title>Login to the Sifaw Control Panel</title></head>
    <body background="images/berriane.jpg">
	<link rel="stylesheet" href="css/style.css" type="text/css">
    <h2 align="center">Login to the Sifaw Control Panel</h2>
    <form method="POST" action="admin.asp">
    <p align="center">User name : <input type="text" name="UserName" size="20"> *<br>
    Password:&nbsp;&nbsp; <input type="password" name="PassWord" size="20"> *</p>
    <p align="center">&nbsp;&nbsp; 
    <input type="submit" value="Login" name="BtnLogin">&nbsp;&nbsp;
    <input type="reset" value="Reset" name="B2"></p>
    <%If Ref = "Try" then
        Response.Write "<i><b>Invalid UserName or Password, try again !!</b></i><hr>"
      end if 'Ref="Try"%>    
    </form>
    </body>
    </html>
  <%else
    If session("Admin") <> "Admin" then 
      Response.Redirect"admin.asp" 
    else 
      if Ref="Log" then%>
      <html>
      <head><title>The Sifaw Control Panel</title></head>
      <body background="images/berriane.jpg">
	  <link rel="stylesheet" href="css/style.css" type="text/css">
      <h1 align="center">The Sifaw Control Panel</h1>
      <p align="center">
      <a href="update.asp?Ref=Main"><b>Mozabite Dictionary</b></a>
      <a href="msgbox.asp?Ref=Main"><b>Guest Book Messages</b></a>
      <a href="dialoguebox.asp?Ref=Main"><b>Dialogue Box</b></a>
      <a href="stats.asp"><b>Statistics</b></a></p>
      <p align="center">
      Home page : <a href="http://www.benbac.fr.st">www.benbac.fr.st</a><br>
      E-Mail : <a href="mailto:webmaster@benbac.fr.st">webmaster@benbac.fr.st</a></p>
      </body>
      </html>
      <%Set Rs = Nothing
      Conn.Close
      Set Conn = Nothing
    end if
  end if
 end if
end if%>
<noframes>