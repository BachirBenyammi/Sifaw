<%If Session("Access")="Site" then%>
<html>
<head><base target="main"></head>
<body>
<link rel="stylesheet" href="css/style.css" type="text/css"> 
<table width="100%">
  <tr>
  <td style="background-color: #000000;"><span class="divise">&nbsp;</span>
  <a href="main.asp" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Main</a>
  <span class="divise">|</span>
  <a href="mzdict.asp" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Mz Dict</a>
  <span class="divise">|</span>
  <a href="guestbook.asp" class="lien" onMouseOver="this.style.color='#FF0000';" onmouseout="this.style.color='#FFFFFF';">Guest Book</a></td>
 </tr>
</table>
</body>
</html>
<%end if%>
<noframes>