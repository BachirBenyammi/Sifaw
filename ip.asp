<% if Session("Access")="Site" then
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

ShortIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
if ShortIP = "" then Short_IP = Request.ServerVariables("REMOTE_ADDR") end if
'ShortIp = "205.521.125.087"
LongIP = NumericIp(ShortIP)

Set conn = Server.CreateObject("ADODB.Connection")
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
SQL = "SELECT Countries.country FROM Ips,Countries WHERE ("&LongIP&" BETWEEN Last_IP AND First_IP) AND (Ips.Country = Countries.Code)"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn
if not rs.eof then Code = rs("Country") end if
rs.close
set rs = nothing
Lien = Request.ServerVariables("HTTP_REFERER")
if Lien = "" then Lien = "Unknoun" end if
Temps = Now
SQL = "INSERT INTO Stat (Ip_Address, Source, Country, Date_Time) Values ('"&ShortIp&"', '"&Lien&"', '"&Code&"', '"&Temps&"')"
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open SQL, conn
Set rs = nothing
conn.close
set conn=nothing
end if%>