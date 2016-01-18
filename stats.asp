<% If session("Admin") <> "Admin" then 
    Response.Redirect"admin.asp" 
  else
   Set conn = Server.CreateObject("ADODB.Connection")
   Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"

    Ref=Request.Querystring("Ref")
    if Ref="" then
    Function NumericIp (ByVal DottedIP)
		Dim i, pos
		Dim PrevPos, num
		If DottedIP = "" Then
			NumericIp = 0
		Else
			For i = 1 To 4
				pos = InStr(PrevPos + 1, DottedIP, ".", 1)
				If i = 4 Then 
					pos = Len(DottedIP) + 1
				End If
				num = Int(Mid(DottedIP, PrevPos + 1, pos - PrevPos - 1))
				PrevPos = pos
				NumericIp = ((num Mod 256) * (256 ^ (4 - i))) +  NumericIp
			Next
		End If
	End Function

SQL = "SELECT * FROM Stat, Countries WHERE Stat.Country = Countries.Code"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn%>
<html>
<head>
<script language="javascript">
  function deletemsg()
  { 
   if(confirmDelete("Are you sure you want to delete this Message(s)"))
  	document.msglist.submit();
  }	
  function confirmDelete(msg)
  {
	if (confirm(msg))
	{
		submitOK = true;
		return true;
	}
	else 
	{
		submitOK = false;
		return false;
	}
  }
  function accSub() 
  {
	if (submitOK == false) 
	{
		submitOK = true;
		return(false);
	} 
	else return(true);
  }	
        </SCRIPT>
</script>
</head>
<body background="images/berriane.jpg">
<link rel="stylesheet" href="css/style.css" type="text/css">
<h1 align="center">Statistics</h1>
<p align=center><A href="admin.asp?Ref=Log">ControlPanel</a>
<A onclick="return (accSub())" href="stats.asp?Ref=Del">Delete All</a></p>
<table border="1" align=center>
<th>Ip Long</th>
<th>Ip Address</th>
<th>Source</th>
<th>Date & Time</th>
<th>Code</th>
<th>Country</th>
<th>Flag</th>
<%do While Not rs.Eof%>
<tr>
<td><%=NumericIp(rs("Ip_Address"))%>&nbsp;</td>
<td><%=rs("Ip_Address")%>&nbsp;</td>
<td><%=rs("source")%>&nbsp;</td>
<td><%=rs("Date_Time")%>&nbsp;</td>
<td><%=rs("Country")%>&nbsp;</td>
<td><%=rs("Code")%>&nbsp;</td>
<td><%="<img src=flags/"&rs("Code")&".png>"%>&nbsp;</td>
</tr>
<%rs.MoveNext
Loop
%>
</table>
<p align=center><A href="admin.asp?Ref=Log">ControlPanel</a>
<A onclick="return (accSub())" href="stats.asp?Ref=Del">Delete All</a></p>
<%rs.Close
set rs=nothing
conn.close
set conn=nothing
else
 if Ref="Del" then
  Set rs = Server.CreateObject("ADODB.Recordset")   
  Rs.Open "Delete * From Stat", Conn
  Response.Redirect "stats.asp"
 end if
 end if
end if%>
<noframes>
</noframes>