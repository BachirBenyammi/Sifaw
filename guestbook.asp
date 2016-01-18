<%  Ref=Request.Querystring("ref")
    if ref="" or ref="try" then%>
      <html>
      <body background="images/berriane.jpg">
	  <!--#include file="menu.asp"-->
	  <link rel="stylesheet" href="css/style.css" type="text/css"> 
	  <h1 align="center">The Sifaw's Guest Book</h1>
	  <p align="center" class="lien">Sign / <a href="guestbook.asp?ref=browse" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Browse</a></p>
	  <%If ref = "try" then
        Response.Write "<i><b>incompelte form or a not valid email address, try again !!</b></i><hr>"
      end if%>    
	  <form action="guestbook.asp?ref=sing" method="POST" name="Form">
      <table width="500" align="center" bgcolor="#C0C0C0" border="#333366">
        <tr> 
          <td valign="top" width="102">Name:</td>
          <td valign="top" width="390">
            <input type="TEXT" size="30" name="UserName"> *</td>
        </tr>
        <tr> 
          <td valign="top" width="102">Email:</td>
          <td valign="top" width="390">
            <input type="TEXT" size="30" name="Email"> *</td>
        </tr>
        <tr> 
          <td valign="top" width="102">Comment:</td>
          <td valign="top" width="390">
            <textarea name="Body" rows="7" cols="46" wrap="Virtual"></textarea>
          </td>
          <td valign="top"> *</td>
        </tr>
        <tr> 
          <td align="center" colspan="2" width="496">
            <input type="submit" value="Sign">
            <input type="reset" value="Reset">
          </td>
        </tr>
      </table>
	  </form>
   	  </body>
	  </html>
	<%else
	Set Conn = Server.CreateObject("ADODB.Connection")      
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
Set Rs = Server.CreateObject("ADODB.Recordset")
	if ref="sing" then
	      UserName=Request.Form("UserName")
	      DlgDateTime=date&" "&time()
	      Email=Request.Form("Email")
	      Body=Request.Form("Body")
          Test=instr(Email,"@")     
          if Test<2 or InStr(Email,".")<5 or len(Email)<7 or UserName="" or Body="" or len(Email)-InStr(Email,".")<1 then
            Response.Redirect"guestbook.asp?ref=try"
	      end if   
	      if InStr(Test,Email,".",1)<Test+3 then
            Response.Redirect"guestbook.asp?ref=try"
	      end if       
          UserName=Replace(UserName,"'","''")
          UserName=Replace(UserName,"<","&lt;")
          UserName=Replace(UserName,">","&gt;")
          Body=Replace(Body,"'","''")
          Body=Replace(Body,"<","&lt;")
          Body=Replace(Body,">","&gt;")
          Body=Replace(Body,VbCrLf,"<br>")                                  
          Rs.Open "Insert Into msgtable (UserName, DlgDateTime, Email, Body) values ('"&UserName&"', '"&DlgDateTime&"', '"&Email&"', '"&Body&"')", conn%>
          <html>
          <body background="images/berriane.jpg">
		  <!--#include file="menu.asp"-->
		  <link rel="stylesheet" href="css/style.css" type="text/css"> 
       	  <h1 align="center">Singing on the Sifaw's Guest Book</h1>
   	      <p align="center" class="lien">Sign 
		  <a href="guestbook.asp?ref=browse" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Browse</a>
		  <a onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';" title="Access to the main page" href="main.asp">Main</a>
		  </p>
          <h4>
          <h1 align="center">Thank you for your comment</h1>
          </body>
          </html>
	    <%else
	      if ref="browse" then
	        Rs.Open "Select * From msgtable Order by msgnu Desc" , Conn%>
            <html>
	        <body background="images/berriane.jpg">		
			<!--#include file="menu.asp"-->
			<link rel="stylesheet" href="css/style.css" type="text/css">
	        <h1 align="center">Browsing The Sifaw's Guest Book</h1>
            <p align="center" class="lien"><a href="guestbook.asp" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Sing</a> / Browse</p>
	        <%do while not Rs.EOF
  	          if Rs("Selected")then %>
                <table align="center" width="500" bgcolor="#C0C0C0" border="#333366">
    	        <tr>
   	              <td colspan="2"><b>Name: </b><%=Rs("UserName")%></td>
	            </tr>  
	            <tr>
   	               <td><b>E-Mail: </b> <a href="mailto:<%=Rs("Email")%>" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';"><%=Rs("Email")%></a></td>
	              <td><b>On: </b><%=Rs("DlgDateTime")%></td>
        	    </tr>
                <tr>
	              <td colspan="2"><b>Comment:</b>
	                <table align="center" width="400"><td><%=Rs("Body")%></td></table>
	              </td>  
	            </tr>
	            </table>
	                <table align="center" width="500"><td align="right"><a href="#top" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Top</a></td></table>
	          <%end if      
	        Rs.movenext
            loop
            Rs.close%>
            <p align="center" class="lien"><a href="guestbook.asp" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Sing</a> / Browse</p>
	        </body>
	        </html>
	      <%end if 'ref="browse"
	    end if 'ref="sing"
	  end if  'ref=""%>
	<noframes>