<%If session("Admin") = "Admin" then %>
<html>
<body>
<link rel="stylesheet" href="css/style.css" type="text/css">
<u><b>Normal Editor :</b></u>
<table border="3" style="width: 100%; border-color: green; background-color: #DFD8DF;">
  <th colspan="5">Normal Editor</th>
  <tr align="center">
    <td width="10%"><input type="Button" value="B" onClick='doShit("Bold")' title="Bold"></td>
    <td width="10%"><input type="Button" value="I" onClick='doShit("Italic")' title="Italic"></td>
    <td width="10%"><input type="Button" value="U" onClick='doShit("Underline")' title="Underline"></td>
    <td width="20%"><input type="Button" value="Text To HTML" onclick="TextToHtml()" title="Text To HTML"></td>
    <td width="20%"><input type="Button" value="HTML To Text" onclick="HtmlToText()" title="HTML To Text"></td>
  </tr>
  <tr><td colspan="5"><iframe width=500 height=150 id=BenEditor></iframe></td></tr>
  <th colspan="5">HTML Editor</th>
  <tr><td colspan="5"><textarea name="BenBrowser" rows="9" cols="61"></textarea></td></tr>
</table>
<script>
  function doShit(shit) 
  {
    var tr = frames.BenEditor.document.selection.createRange()
    tr.execCommand(shit)
    tr.select()
    frames.BenEditor.focus()
  }
  function TextToHtml() 
  {
    BenBrowser.value = BenEditor.document.body.innerHTML;
  }
  function HtmlToText() 
  {
    BenEditor.document.body.innerHTML= BenBrowser.value;
  }
  BenEditor.document.designMode = "on" 
</script>
<%end if%>
<noframes>