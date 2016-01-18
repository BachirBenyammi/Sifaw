<%If session("Admin") <> "Admin" then 
    Response.Redirect"admin.asp" 
  else
  Set Conn = Server.CreateObject("ADODB.Connection")      
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
Set Rs = Server.CreateObject("ADODB.Recordset")
  Ref=Request.Querystring("Ref")
    if Ref="" then
      Response.Redirect "admin.asp"
    else  
    if Ref="Main" then
        Rs.Open "Select * From msgtable" , Conn%>
        <html>
        <head>
        <title>Guest Book Messages</title>
        <SCRIPT language=javascript1.2>
          var submitOK = true;
          function SelectAll() 
          {
          	len = document.msglist.elements.length;
        	var i=0;
        	while(i!=len)
            {
              if (document.msglist.elements[i].type=='checkbox') 
                if (document.msglist.elements[i].name=='CB')
            	  document.msglist.elements[i].checked = !document.msglist.elements[i].checked;
  		      i++;
        	}
          }
  function deletemsg()
  { 
 	len = document.msglist.elements.length;
	var i=0;
	var j=0;
	while(i!=len)
	{
		if (document.msglist.elements[i].type=='checkbox') 
    	  if (document.msglist.elements[i].name=='CB')
          if (document.msglist.elements[i].checked ==true )
    	 	j++;
		i++;
	}
    if (j==0)
       alert(document.msglist.DELETEWHAT.value);
    else   
    { 
      if(confirmDelete(document.msglist.DELETECONFIRM.value))
      {
       	document.msglist.DELETE.value = "1";
       	document.msglist.submit();
      } 	
      else
	    document.msglist.DELETE.value = "0";  
    }	            
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
        </head>
        <body  background="images/berriane.jpg">
		<link rel="stylesheet" href="css/style.css" type="text/css">
        <h1 align="center">The Guest Book Messages</h1>
        <form name=msglist onsubmit="return accSub();" method=post action="msgbox.asp?Ref=Delete">
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())" href="javascript:deletemsg()">Delete</A>
        <A href="admin.asp?Ref=Log">ControlPanel</a>
        <table border="1" cellspacing="0" width="100%">
        <th width="10%">N°</th>
        <th width="5%">Selected</th>        
        <th width="30%">UserName</th>
        <th width="35%">Email</th>        
        <th width="20%">Date/Time</th>
        <%do while not Rs.EOF%> 
        <tr align="center">
        <td width="10%"><input type="checkbox" name="CB" value=<%=Rs("msgnu")%>>
        <a href="msgbox.asp?Ref=Read&reference=<%=Rs("msgnu")%>"><%=Rs("msgnu")%></a></td>
        <td width="5%">
        <% if Rs("Selected") then 
             Response.Write("Yes")
           else
             Response.Write("No")
           end if%>&nbsp;</td>     
        <td width="30%"><%=Rs("UserName")%>&nbsp;</td>
        <td width="35%"><a href="mailto:<%=Rs("Email")%>"><%=Rs("Email")%></a>&nbsp;</td>
        <td width="20%"><%=Rs("DlgDateTime")%>&nbsp;</td>
        </tr>          
        <%Rs.movenext
        loop%>
        </table>
        <INPUT type=hidden value="Please select the Message(s) you want to delete !!" name=DELETEWHAT> 
        <INPUT type=hidden value="Are you sure you want to delete the selected Message(s) ?" name=DELETECONFIRM> 
        <INPUT type=hidden value=0 name=DELETE> 
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())" href="javascript:deletemsg()">Delete</A>
        <A href="admin.asp?Ref=Log">ControlPanel</a></form>
        </body>
        </html>
      <%Rs.close
      else
        if Ref="Delete" then
          CB=Request.Form("CB")
          msgref=Request.Querystring("reference")
          if msgref="" then
            Rs.Open "Delete * From msgtable WHERE msgnu in (Select msgnu from msgtable WHERE instr('"&CB&"',msgnu))", Conn
          else 
            msgref=CINT(msgref)
            Rs.Open "Delete * From msgtable WHERE msgnu ="&msgref, Conn
          end if 'msgref=""
          Response.Redirect"msgbox.asp?Ref=Main"
        else 
          if Ref="Read" then
            msgref=Request.Querystring("reference")
            msgref=CINT(msgref)
            Rs.Open "Select * from msgtable WHERE msgnu=" & msgref , Conn%>
            <html>
            <head>
            <title>Read a Guest Book Message</title>
            <SCRIPT language=javascript1.2>
            var submitVarRead = true;    
            function deletemsg()
            {
              if(confirmDelete(document.msgread.DELETECONFIRM.value))
		      {
       	        document.msgread.DELETE.value = "1";
            	document.msgread.submit();
		      } 	
            }
            function confirmDelete(msg)
			{
			  if (confirm(msg))
			  {
				submitVarRead = true;
				return true;
			  }
			  else 
			  {
				submitVarRead = false;
				return false;
			  }
			}
   		    function SubmitRead()
	        { 
      	      document.msgread.submit();    
            }    
    	    function ResetRead()
            { 
              document.msgread.reset();    
            }
            function accSubRead() 
            { 
              if (submitVarRead == false) 
           	  {
	    	    submitVarRead = true;
        	    return(false);
    	      } 
        	  else 
        	    return(true);
		    }     
            </SCRIPT>
          </head>
          <body  background="images/berriane.jpg">
		  <link rel="stylesheet" href="css/style.css" type="text/css">
          <form name="frm" method="POST" action="msgbox.asp?Ref=Sel&reference=<%=Rs("msgnu")%>">
          <h1 align="center">A Guest Book Message N° <%=msgref%></h1>
          <p align="center"><b>Activate :</b>
          <select name="Selected">
          <% if Rs("Selected")then 
               Response.Write("<option Selected>Yes</option><option>No</option>")
             else 
               Response.Write("<option>Yes</option><option Selected>No</option>")
             end if%>
          </select>
          <input type="Submit" value="Update"><p>
          </form>
          <form name="msgread" onsubmit="return accSubRead();" method="POST" action="msgbox.asp?Ref=Delete&reference=<%=Rs("msgnu")%>">
          <P align="center">
          <A onclick="return (accSubRead())" href="javascript:deletemsg()">Delete</A>
          <A href="msgbox.asp?Ref=Main">Main</A>
          <A href="admin.asp?Ref=Log">ControlPanel</a><P align="left">
          <input type="hidden" Name="msgref" Value="<%=msgref%>">
          <table align="center" border="0">
          <th colspan="2" align=center>Message</th>
          <tr><td><b>User name :</b></td>
          <td><input type="text" name="UserName" size="35" value="<%=Rs("UserName")%>" READONLY></td></tr>
          <tr><td><b>Email :</b></td>
          <td><input type="text" name="Email" size="36" value="<%=Rs("Email")%>" READONLY></td></tr>
          <tr><td><b>Date/Time: </b></td><td><input type="text" name="DlgDateTime" size="36" value="<%=Rs("dlgdatetime")%>" READONLY></td></tr>
          <tr><td valign=top><b>Body:</b></td>
          <td><textarea rows="16" name="Body" cols="64" READONLY><%=Rs("Body")%></textarea></td></tr>
          </table>
          <p align="center">
          <A onclick="return (accSubRead())"href="javascript:deletemsg()">Delete</A>
          <A href="msgbox.asp?Ref=Main">Main</A>
          <A href="admin.asp?Ref=Log">ControlPanel</a>
          <INPUT type=hidden value="Are you sure you want to delete the selected Message(s) ?" name=DELETECONFIRM> 
          <INPUT type=hidden value=0 name=DELETE>         
          </form>
          <table>         
          </body>
          </html>
          <%Rs.Close
         else
           if Ref="Sel" Then
             MsgReff=Request.Querystring("reference")
             MsgRef=CINT(MsgReff) 
             Select Case Request.Form("Selected")
               Case "Yes" Rs.Open "Update msgtable Set Selected=TRUE WHERE msgnu=" & MsgRef, Conn
               Case "No" Rs.Open "Update msgtable Set Selected=FALSE WHERE msgnu=" & MsgRef, Conn
             end Select
             Response.Redirect"msgbox.asp?Ref=Main"
           end if 'Ref=Sel" 
         end if 'Ref=Read"   
       end if 'Ref="Delete"
     end if 'Ref="Main"
    end if 'ref=""        
     Set Rs = Nothing
     Conn.Close
     Set Conn = Nothing
  end if 'Entre<>Ok%><noframes></noframes>