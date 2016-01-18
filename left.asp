<% if Session("Access")="Site" then
Set Conn = Server.CreateObject("ADODB.Connection")      
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
Set Rs = Server.CreateObject("ADODB.Recordset")
Rs.Open "Select visitors_number From visitorstable" , Conn
 if not Rs.EOF then
  newNum=Rs("visitors_number")+1
end if
Rs.Close
Rs.Open "Update visitorstable Set visitors_number='"&newNum&"'", Conn
Set Rs = Nothing
Conn.Close
Set Conn = Nothing%>
<html>
<head><base target="main">
<script language="Javascript" src="js/js.js"></script>
</head>
<body background="images/background.jpg" onLoad="debuteDate();debuteTemps()" onUnload="clearTimeout(ddate);clearTimeout(ttime)">
<link rel="stylesheet" href="css/style.css" type="text/css">
<script language="JavaScript1.2">
function AddToFovorite(){
    window.external.AddFavorite("http://www.sifaw.fr.st","Sifaw");
}
</script>
<script language="JavaScript"><!--
function jumpto(){
  window.open(document.frm.list.options[document.frm.list.selectedIndex].value,target="main");
}// -->
</script>
<br><center><a href="http://www.sifaw.fr.st" target="_parent" ><img alt="http://www.sifaw.fr.st" src="images/logo.gif" width="120" height="50"></a></center>
<p>
<a class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';" title="Access to the main page" href="main.asp">Main</a><br>
<a href="mzdict.asp" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';" title="Access to the Mozabite Dictionary">Mz Dict</a><br>
<a href="guestbook.asp" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';" title="Access to the Guest Book">Guest Book</a></p>
<p><a href class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';" onClick="this.style.behavior='url(#default#homepage)';this.setHomePage('http://www.sifaw.fr.st');">Set as Default <br>Home Page</a><br><br>
<a href onclick="javascript:AddToFovorite();" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Add to Favorites</a></p>
<p class="lien">Counter = <%=newNum%></p>
<form name="frm">
<p><span class="lien">Go to :</span><br>
<select name="list">
  <option selected value="main.asp">Main</option>
  <option value="mzdict.asp">Mz Dict</option>
  <option value="guestbook.asp">Guest Book</option>  
 </select>
<input type="button" value="Go!" onClick="jumpto()">
 </p>
 </form>
<p class="lien"><span id="heure"></span><br><span id="jour"></span></p>
</body>
</html>
<%end if%>
<noframes>