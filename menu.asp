<%If Session("Access")="Site" then%>
<link rel="stylesheet" type="text/css" href="css/menu.css">
<div class="skin0" id="ie5menu" onmouseover="highlightie5()" onclick="jumptoie5()" onmouseout="lowlightie5()">
  <div class="menuitems" url="javascript:history.go(-1)"><img src="images/back.gif" width="12" height="12">&nbsp;Back</div>
  <div class="menuitems" url="javascript:history.go(1)"><img src="images/next.gif" width="12" height="12">&nbsp;Next</div><hr>
  <div class="menuitems" url="main.asp"><img src="images/main.gif" width="12" height="12">&nbsp;Main</div>
  <div class="menuitems" url="mzdict/default.asp"><img src="images/book.gif" width="12" height="12">&nbsp;Mz Dict</div>
  <div class="menuitems" url="guestbook.asp"><img src="images/guestbook.gif" width="12" height="12">&nbsp;Guest Book</div>
</div>
<script language="Javascript" src="js/menu.js"></script>
<%end if%>