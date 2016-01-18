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
        Rs.Open "Select * From dialoguetable" , Conn%>
        <html>
        <head>
        <title>A Dialogue Box</title>
        <SCRIPT language=javascript1.2>
          var submitOK = true;
          function SelectAll() 
          {
          	len = document.dialoguelist.elements.length;
        	var i=0;
        	while(i!=len)
            {
              if (document.dialoguelist.elements[i].type=='checkbox') 
            	  document.dialoguelist.elements[i].checked = !document.dialoguelist.elements[i].checked;
  		      i++;
        	}
          }
  function deletedialogue()
  { 
 	len = document.dialoguelist.elements.length;
	var i=0;
	var j=0;
	while(i!=len)
	{
		if (document.dialoguelist.elements[i].type=='checkbox') 
          if (document.dialoguelist.elements[i].checked ==true )
    	 	j++;
		i++;
	}
    if (j==0)
       alert(document.dialoguelist.DELETEWHAT.value);
    else   
    { 
      if(confirmDelete(document.dialoguelist.DELETECONFIRM.value))
      {
       	document.dialoguelist.DELETE.value = "1";
       	document.dialoguelist.submit();
      } 	
      else
	    document.dialoguelist.DELETE.value = "0";  
    }	            
  }
  function confirmDelete(dialogue)
  {
	if (confirm(dialogue))
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
        <h1 align="center">A Dialogue Box</h1>
        <form name=dialoguelist onsubmit="return accSub();" method=post action="dialoguebox.asp?Ref=Delete">
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())" href="javascript:deletedialogue()">Delete</A>
        <A href="dialoguebox.asp?Ref=Add">Add</A>
        <A href="admin.asp?Ref=Log">ControlPanel</a>
        <table border="1" cellspacing="0" width="100%">
        <th width="20%">N°</th>
        <th width="40%">UserName</th>
        <th width="40%">Date/Time</th>
        <%do while not Rs.EOF%> 
        <tr>
        <td width="20%"><input type="checkbox" name="CB" value=<%=Rs("dialoguenu")%>>
        <a href="dialoguebox.asp?Ref=Read&reference=<%=Rs("dialoguenu")%>"><%=Rs("dialoguenu")%></a></td>
        <td width="40%"><%=Rs("UserName")%>&nbsp;</td>
        <td width="40%"><%=Rs("dlgdatetime")%>&nbsp;</td>
        </tr>
        <%Rs.movenext
        loop%>
        </table>
        <INPUT type=hidden value="Please select the Dialogue(s) you want to delete !!" name=DELETEWHAT> 
        <INPUT type=hidden value="Are you sure you want to delete the selected Dialogue(s) ?" name=DELETECONFIRM> 
        <INPUT type=hidden value=0 name=DELETE> 
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())"href="javascript:deletedialogue()">Delete</A>
        <A href="dialoguebox.asp?Ref=Add">Add</A>
        <A href="admin.asp?Ref=Log">ControlPanel</a></form>
        </body>
        </html>
      <%Rs.close
      else
        if Ref="Delete" then
          CB=Request.Form("CB")
          dialogueref=Request.Querystring("reference")
          if dialogueref="" then
            Rs.Open "Delete * From dialoguetable WHERE dialoguenu in (Select dialoguenu from dialoguetable WHERE instr('"&CB&"',dialoguenu))", Conn
          else 
            dialogueref=CINT(dialogueref)
            Rs.Open "Delete * From dialoguetable WHERE dialoguenu ="&dialogueref, Conn
          end if 'dialogueref=""
          Response.Redirect"dialoguebox.asp?Ref=Main"
        else 
          if Ref="Read" then
            dialogueref=Request.Querystring("reference")
            dialogueref=CINT(dialogueref)
            Rs.Open "Select * from dialoguetable WHERE dialoguenu=" & dialogueref , Conn%>
            <html>
            <head>
            <title>Read a Dialogue</title>
            <SCRIPT language=javascript1.2>
            var submitVarRead = true;    
            function deletedialogue()
            {
              if(confirmDelete(document.dialogueread.DELETECONFIRM.value))
		      {
       	        document.dialogueread.DELETE.value = "1";
            	document.dialogueread.submit();
		      } 	
            }
            function confirmDelete(dialogue)
			{
			  if (confirm(dialogue))
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
      	      document.dialogueread.submit();    
            }    
    	    function ResetRead()
            { 
              document.dialogueread.reset();    
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
          <h2 align="center">A Dialogue N° <%=dialogueref%> <%=dlgdatetime%></h2>
          <form name="dialogueread" onsubmit="return accSubRead();" method="POST" action="dialoguebox.asp?Ref=Delete&reference=<%=Rs("dialoguenu")%>">
          <P align="center">
          <A onclick="return (accSubRead())"href="javascript:deletedialogue()">Delete</A>
          <A href="dialoguebox.asp?Ref=Main">Main</A>
          <A href="admin.asp?Ref=Log">ControlPanel</a>
          <input type="hidden" Name="dialogueref" Value="<%=dialogueref%>"><br><br>
          <table align="center" border="3">
		  <tr>
		   <td colspan="2"><b>User name:</b></td>
		   <td><input type="text" name="UserName" size="20" value="<%=Rs("UserName")%>" READONLY></td>
		  </tr> 
          <tr>
		   <td colspan="2"><b>Date/Time:</b></td>
		   <td><input type="text" name="DlgDateTime" size="20" value="<%=Rs("dlgdatetime")%>" READONLY></td>
		  </tr> 
          <tr><td colspan="3"><b>Body:</b><br><br>
          <textarea rows="16" name="Body" cols="64" READONLY><%=Rs("Body")%></textarea></td></tr>
		  </table>
          <p align="center">
          <A onclick="return (accSubRead())"href="javascript:deletedialogue()">Delete</A>
          <A href="dialoguebox.asp?Ref=Main">Main</A>
          <A href="admin.asp?Ref=Log">ControlPanel</a>
          <INPUT type=hidden value="Are you sure you want to delete the selected Dialogue ?" name=DELETECONFIRM> 
          <INPUT type=hidden value=0 name=DELETE>         
          </form>
          </body>
          </html>
          <%Rs.Close
          else
            if Ref="Add" then%>
              <html>
              <head>
              <title>A New Dialogue</title>
              <SCRIPT language=javascript1.2>         
      var submitVarNew = true;
      function SubmitNew()
      { 
        document.dialoguenew.submit();    
      }      
      function ResetNew()
      { 
        document.dialoguenew.reset();    
      }
      function accSubNew() 
      { 
    	if (submitVarNew == false) 
     	{
	    	submitVarNew = true;
		    return(false);
    	} 
    	else return(true);
      }
                </SCRIPT>
                </head>
                <body  background="images/berriane.jpg">
				<link rel="stylesheet" href="css/style.css" type="text/css">
                <h2 align="center">A New Dialogue</h2>
                <form name="dialoguenew" onsubmit="return accSubNew();" method="POST" action="dialoguebox.asp?Ref=New">
                <p align="center"><a href="javascript:SubmitNew()">Save</a> <A href="javascript:ResetNew()">Reset</A>
                <A href="dialoguebox.asp?Ref=Main">Main</A>
                <A href="admin.asp?Ref=Log">ControlPanel</a></p>
                <table align="center" border="3">
				<tr>
                 <td colspan="2"><b>User name:</b></td>
				 <td><input type="text" name="UserName" size="20"></td>
                </tr>
				<tr><td colspan="3"><b>Body:</b><br><br>
				<textarea rows="16" name="Body" cols="64"></textarea></td></tr>
				</table>
                <p align="center"><a href="javascript:SubmitNew()">Save</a> <A href="javascript:ResetNew()">Reset</A>
                <A href="dialoguebox.asp?Ref=Main">Main</A>
                <A href="admin.asp?Ref=Log">ControlPanel</a></form>
                </body>
                </html>
            <%else 
                if Ref="New" then 
                  UserName=Request.Form("UserName")
                  DlgDateTime=date()&" "&time()
                  Body=Request.Form("Body")
                  Body=Replace(Body,"'","''")
                  UserName=Replace(UserName,"'","''")
                  Rs.Open "Insert Into dialoguetable ( username, dlgdatetime, body) Values (' " & UserName& " ',' " & DlgDateTime& " ',' " & Body& " ')" , Conn
                  Response.Redirect"dialoguebox.asp?Ref=Main" 
                end if 'Ref="New"   
              end if 'Ref="Add"     
          end if 'Ref=Read"   
        end if 'Ref="Delete"
      end if 'Ref="Main"
      end if 'ref=""         
      Set Rs = Nothing
      Conn.Close
      Set Conn = Nothing
    end if 'Entre<>Ok%>
	<noframes>