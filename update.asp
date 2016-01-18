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
        Rs.Open "Select * From mztable" , Conn%>
        <html>
        <head>
        <title>Mozabite Dictionary</title>
        <SCRIPT language=javascript1.2>
          var submitOK = true;
          function SelectAll() 
          {
          	len = document.WordList.elements.length;
        	var i=0;
        	while(i!=len)
            {
              if (document.WordList.elements[i].type=='checkbox') 
          	  document.WordList.elements[i].checked = !document.WordList.elements[i].checked;
  		      i++;
        	}
          }
  function deleteMsg()
  { 
 	len = document.WordList.elements.length;
	var i=0;
	var j=0;
	while(i!=len)
	{
		if (document.WordList.elements[i].type=='checkbox') 
          if (document.WordList.elements[i].checked ==true )
    	 	j++;
		i++;
	}
    if (j==0)
       alert(document.WordList.DELETEWHAT.value);
    else   
    { 
      if(confirmDelete(document.WordList.DELETECONFIRM.value))
      {
       	document.WordList.DELETE.value = "1";
       	document.WordList.submit();
      } 	
      else
	    document.WordList.DELETE.value = "0";  
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
		<h1 align="center">The Mozabite Dictionary</h1>
        <form method=post action="update.asp?Ref=Search">
        <p align="center">Search :                          
        <input type="text" name="SearchText" size="20"> Language :
        <select size="1" name="Lang">
          <option selected>En</option>
          <option>Fr</option>
          <option>Ar</option>
          <option>MzLa</option>
          <option>MzAr</option>
        </select>
        <input type="submit" value="Search"></p>
        </form>
        <form name=WordList onsubmit="return accSub();" method=post action="update.asp?Ref=Delete">  
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())" href="javascript:deleteMsg()">Delete</A>
        <A href="update.asp?Ref=Add">Add</A>
		<A href="admin.asp?Ref=Log">ControlPanel</a>
        <table border="1" width="100%" cellspacing="0">    
        <th width="10%">N°</th>     
        <th width="18%">En</th>     
        <th width="18%">Fr</th>     
        <th width="18%">Ar</th>     
        <th width="18%">MzLa</th>     
        <th width="18%">MzAr</th>    
      <%do while not Rs.EOF%>  
	    <tr>  
        <td width="10%" align="left">
        <input type="checkbox" name="CB" value=<%=Rs("wordnu")%>>
        <a href="update.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("wordnu")%></a>
        </td>
        <td width="18%" align="center"><%=Rs("entrans")%>&nbsp;</td>
        <td width="18%" align="center"><%=Rs("frtrans")%>&nbsp;</td>
        <td width="18%" align="center"><%=Rs("artrans")%>&nbsp;</td>
        <td width="18%" align="center"><%=Rs("mzlatrans")%>&nbsp;</td>
        <td width="18%" align="center"><%=Rs("mzartrans")%>&nbsp;</td>
        </tr>
      <%Rs.movenext
        loop%>
        </table>
        <INPUT type=hidden value="Please select the word(s) you want to delete !!" name=DELETEWHAT> 
        <INPUT type=hidden value="Are you sure you want to delete the selected word(s) ?" name=DELETECONFIRM> 
        <INPUT type=hidden value=0 name=DELETE> 
        <A href="javascript:SelectAll()">All/None</A>
        <A onclick="return (accSub())" href="javascript:deleteMsg()">Delete</A>
        <A href="update.asp?Ref=Add">Add</A>
		<A href="admin.asp?Ref=Log">ControlPanel</a>
        </form>
        <form method=post action="update.asp?Ref=Search">
        <P align="center">Search :                          
        <input type="text" name="SearchText" size="20"> Language :
        <select size="1" name="Lang">
          <option selected>En</option>
          <option>Fr</option>
          <option>Ar</option>
          <option>MzLa</option>
          <option>MzAr</option>
        </select>
        <input type="submit" value="Search"></p>
        </form>
        </body>
        </html>
      <%Rs.close
      else
        if Ref="Delete" then
          CB=Request.Form("CB")
          WordRef=Request.Querystring("reference")
          if WordRef="" then
            Rs.Open "Delete * From mztable WHERE wordnu in (Select wordnu from mztable WHERE instr('"&CB&"',wordnu))", Conn
          else 
            WordRef=CINT(WordRef)
            Rs.Open "Delete * From mztable WHERE wordnu ="&WordRef, Conn
          end if 'WordRef=""
          Response.Redirect"update.asp?Ref=Main"
        else 
          if Ref="Read" then
            WordRef=Request.Querystring("reference")
            WordRef=CINT(WordRef)
            Rs.Open "Select * from mztable WHERE wordnu=" & WordRef , Conn%>
            <html>
            <head>
            <title>Mozabite Dictionary !!</title>
            <SCRIPT language=javascript1.2>
            function Editor()
            {
              window.open("beneditor.asp","","height=400,width=550,top=50,left=50,scrollbars=yes,resizable=yes")
            }
            function ViewThis(code)
            {
              Win = window.open("","","height=200,width=400,top=50,left=50,scrollbars=yes,resizable=yes");
              Win.document.write('<html><head><title>HTML Browser</title></head><body>');
              Win.document.write(code+ '</body></html>');
              Win.document.close();
            }            
            function ViewAll(code)
            {
              Win = window.open("","","height=400,width=600,top=50,left=50,scrollbars=yes,resizable=yes")
              Win.document.write('<html><head><title>HTML Browser</title></head><body>');
              Win.document.write('EnTrans : ' + code.EnTrans.value + '<br>');
              Win.document.write('FrTrans : ' + code.FrTrans.value + '<br>');
              Win.document.write('ArTrans : ' + code.ArTrans.value + '<br>');
              Win.document.write('MzLaTrans : ' + code.MzLaTrans.value + '<br>');
              Win.document.write('MzArTrans : ' + code.MzArTrans.value + '<br>');
              Win.document.write('<p>EnExp :<br>' + code.EnExp.value + '</p>');
              Win.document.write('<p>FrExp :<br>' + code.FrExp.value + '</p>');
              Win.document.write('<p>ArExp :<br>' + code.ArExp.value + '</p>');
              Win.document.write('<p>MzLaExp :<br>' + code.MzLaExp.value + '</p>');
              Win.document.write('<p>MzArExp :<br>' + code.MzArExp.value + '</p>');                  
              Win.document.write('</body></html>');
              Win.document.close();
            }
            var submitVarModif = true;  
            function ModifMsg()
            {
	          if(confirmModif(document.WordModif.MODIFCONFIRM.value))
		      {
       			document.WordModif.MODIF.value = "1";
		       	document.WordModif.submit();
		      } 	
		      else
	    		document.WordModif.MODIF.value = "0";  
		    }	            
		    function confirmModif(msg)
			{
		   	  if (confirm(msg))
			  {
				submitVarModif = true;
				return true;
			  }
			  else 
			  {
				submitVarModif = false;
				return false;
			  }
		    }            
      function ResetModif()
      { 
        document.WordModif.reset();    
      }
     function accSubModif() 
     { 
    	if (submitVarModif == false) 
     	{
	    	submitVarModif = true;
		    return(false);
    	} 
    	else return(true);
     }     
            </SCRIPT>
          </head>
          <body  background="images/berriane.jpg">
		  <link rel="stylesheet" href="css/style.css" type="text/css">
          <h1 align="center">Word N° <%=WordRef%></h1>
          <form name="WordModif" onsubmit="return accSubModif();" method="POST" action="update.asp?Ref=Modif&reference=<%=Rs("wordnu")%>">
          <P align="center">
          <A onclick="return (accSubModif())"href="javascript:ModifMsg()">Modif</A>
          <A href="javascript:ResetModif()">Reset</A>
          <A href="update.asp?Ref=Delete&reference=<%=Rs("wordnu")%>">Delete</A>
          <A href="update.asp?Ref=Main">Main</A>
          <a href="javascript:ViewAll(document.WordModif);">View</a> 
          <a href="javascript:Editor();">Editor</a>
		  <A href="admin.asp?Ref=Log">ControlPanel</a></P>
          <input type="hidden" Name="WordRef" Value="<%=WordRef%>">
          <table width="80%" align="center"><td>
          En Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="text" name="EnTrans" size="20" value="<%=Rs("EnTrans")%>"><br><br>
          Fr Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="text" name="FrTrans" size="20" value="<%=Rs("FrTrans")%>"><br><br>
          Ar Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
          <input type="text" name="ArTrans" size="20" value="<%=Rs("ArTrans")%>"><br><br>
          Mz La Trans :<input type="text" name="MzLaTrans" size="20" value="<%=Rs("MzLaTrans")%>"><br><br>
          Mz Ar Trans :<input type="text" name="MzArTrans" size="20" value="<%=Rs("MzArTrans")%>"><br><br>
          En Exp : <a href="javascript:Editor();">Editor</a>
          <a href="javascript:ViewThis(document.WordModif.EnExp.value);">View</a><br>
          <textarea name="EnExp" rows="5" cols="50"><%=Rs("EnExp")%></textarea><br><br>
          Fr Exp :<a href="javascript:Editor();">Editor</a>
           <a href="javascript:ViewThis(document.WordModif.FrExp.value);">View</a><br>
          <textarea name="FrExp" rows="5" cols="50"><%=Rs("FrExp")%></textarea><br><br>
          Ar Exp : <a href="javascript:Editor();">Editor</a>
          <a href="javascript:ViewThis(document.WordModif.ArExp.value);">View</a><br>
          <textarea name="ArExp" rows="5" cols="50" ><%=Rs("ArExp")%></textarea><br><br>
          Mz La Exp : <a href="javascript:Editor();">Editor</a>
          <a href="javascript:ViewThis(document.WordModif.MzLaExp.value);">View</a><br>
          <textarea name="MzLaExp" rows="5" cols="50"><%=Rs("MzLaExp")%></textarea><br><br>
          Mz Ar Exp : <a href="javascript:Editor();">Editor</a>
          <a href="javascript:ViewThis(document.WordModif.MzArExp.value);">View</a><br>
          <textarea name="MzArExp" rows="5" cols="50"><%=Rs("MzArExp")%></textarea><br><br>
          </td></table>
          <INPUT type=hidden value="Are you sure you want to modif the selected word ?" name=MODIFCONFIRM> 
          <INPUT type=hidden value=0 name=MODIF>           
          <p align="center">
          <A onclick="return (accSubModif())"href="javascript:ModifMsg()">Modif</A>
          <A href="javascript:ResetModif()">Reset</A>
          <A href="update.asp?Ref=Delete&reference=<%=Rs("wordnu")%>">Delete</A> 
          <A href="update.asp?Ref=Main">Main</A>
          <a href="javascript:ViewAll(document.WordModif);">View</a> 
          <a href="javascript:Editor();">Editor</a>
		  <A href="admin.asp?Ref=Log">ControlPanel</a></p>    
          </form>
          </body>
          </html>
          <%Rs.Close
          else
            if Ref="Modif" then
              WordRef=Request.Querystring("reference")
              WordRef=CINT(WordRef)
              EnTrans=Request.Form("EnTrans")
              FrTrans=Request.Form("FrTrans")
              ArTrans=Request.Form("ArTrans")
              MzLaTrans=Request.Form("MzLaTrans")
              MzArTrans=Request.Form("MzArTrans")
              EnExp=Request.Form("EnExp")
              FrExp=Request.Form("FrExp")
              ArExp=Request.Form("ArExp")
              MzLaExp=Request.Form("MzLaExp")
              MzArExp=Request.Form("MzArExp")
              EnTrans=Replace(EnTrans,"'","''")
              FrTrans=Replace(FrTrans,"'","''")
              ArTrans=Replace(ArTrans,"'","''")
              MzLaTrans=Replace(MzLaTrans,"'","''")
              MzArTrans=Replace(MzArTrans,"'","''")               
              Rs.Open "Update mztable Set entrans='"&EnTrans&"',frtrans='"&FrTrans&"',artrans='"&ArTrans&"',mzlatrans='"&MzLaTrans&"',mzartrans='"&MzArTrans&"',enexp='"&EnExp&"',frexp='"&FrExp&"',arexp='"&ArExp&"',mzlaexp='"&MzLaExp&"',mzarexp='"&MzArExp&"' WHERE wordnu=" & WordRef, Conn
              Response.Redirect"update.asp?Ref=Main"
            else
              if Ref="Add" then%>
                <html>
                <head>
                <title>A New word !!</title>
                <SCRIPT language=javascript1.2>
                function Editor()
                {
                  window.open("beneditor.asp","","height=400,width=550,top=50,left=50,scrollbars=yes,resizable=yes")
                }
                function ViewThis(code)
                {
                  Win = window.open("","","height=200,width=400,top=50,left=50,scrollbars=yes,resizable=yes");
                  Win.document.write('<html><head><title>HTML Browser</title></head><body>');
                  Win.document.write(code+ '</body></html>');
                  Win.document.close();
                }                
                function ViewAll(code)
                {
                  Win = window.open("","","height=400,width=600,top=50,left=50,scrollbars=yes,resizable=yes")
                  Win.document.write('<html><head><title>HTML Browser</title></head><body>');
                  Win.document.write('EnTrans : ' + code.EnTrans.value + '<br>');
                  Win.document.write('FrTrans : ' + code.FrTrans.value + '<br>');
                  Win.document.write('ArTrans : ' + code.ArTrans.value + '<br>');
                  Win.document.write('MzLaTrans : ' + code.MzLaTrans.value + '<br>');
                  Win.document.write('MzArTrans : ' + code.MzArTrans.value + '<br>');
                  Win.document.write('<p>EnExp :<br>' + code.EnExp.value + '</p>');
                  Win.document.write('<p>FrExp :<br>' + code.FrExp.value + '</p>');
                  Win.document.write('<p>ArExp :<br>' + code.ArExp.value + '</p>');
                  Win.document.write('<p>MzLaExp :<br>' + code.MzLaExp.value + '</p>');
                  Win.document.write('<p>MzArExp :<br>' + code.MzArExp.value + '</p>');                  
                  Win.document.write('</body></html>');
                  Win.document.close();
                 }
      var submitVarNew = true;
      function SubmitNew()
      { 
        document.WordNew.submit();    
      }      
      function ResetNew()
      { 
        document.WordNew.reset();    
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
                <h1 align="center">A New word</h1>
                <form name="WordNew" onsubmit="return accSubNew();" method="POST" action="update.asp?Ref=New">
                <p align="center">
                <a href="javascript:SubmitNew()">Save</a> <A href="javascript:ResetNew()">Reset</A>
                <A href="update.asp?Ref=Main">Main</A>
                <a href="javascript:ViewAll(document.WordNew);">View</a> 
                <a href="javascript:Editor();">Editor</a>
				<A href="admin.asp?Ref=Log">ControlPanel</a></p>
                <table width="80%" align="center"><td>
                En Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input type="text" name="EnTrans" size="20" ><br><br>
                Fr Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input type="text" name="FrTrans" size="20" ><br><br>
                Ar Trans :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                <input type="text" name="ArTrans" size="20"><br><br>
                Mz La Trans :<input type="text" name="MzLaTrans" size="20" ><br><br>
                Mz Ar Trans :<input type="text" name="MzArTrans" size="20"><br><br>
                En Exp : <a href="javascript:Editor();">Editor</a>
                <a href="javascript:ViewThis(document.WordNew.EnExp.value);">View</a><br>
                <textarea name="EnExp" rows="5" cols="50"> </textarea><br><br>
                Fr Exp : <a href="javascript:Editor();">Editor</a>
                <a href="javascript:ViewThis(document.WordNew.FrExp.value);">View</a><br>
                <textarea name="FrExp" rows="5" cols="50"> </textarea><br><br>
                Ar Exp : <a href="javascript:Editor();">Editor</a>
                <a href="javascript:ViewThis(document.WordNew.ArExp.value);">View</a><br>
                <textarea name="ArExp" rows="5" cols="50"> </textarea><br><br>
                Mz La Exp : <a href="javascript:Editor();">Editor</a>
                <a href="javascript:ViewThis(document.WordNew.MzLaExp.value);">View</a><br>
                <textarea name="MzLaExp" rows="5" cols="50"> </textarea><br><br>
                Mz Ar Exp : <a href="javascript:Editor();">Editor</a>
                <a href="javascript:ViewThis(document.WordNew.MzArExp.value);">View</a><br>
                <textarea name="MzArExp" rows="5" cols="50"> </textarea><br>
                </td></table>
                <p align="center">
                <a href="javascript:SubmitNew()">Save</a> <A href="javascript:ResetNew()">Reset</A>
                <A href="update.asp?Ref=Main">Main</A> 
                <a href="javascript:ViewAll(document.WordNew);">View</a> 
                <a href="javascript:Editor();">Editor</a>
				<A href="admin.asp?Ref=Log">ControlPanel</a></p>
                </form>
                </body>
                </html>
            <%else 
                if Ref="New" then 
                  EnTrans=Request.Form("EnTrans")
                  FrTrans=Request.Form("FrTrans")
                  ArTrans=Request.Form("ArTrans")
                  MzLaTrans=Request.Form("MzLaTrans")
                  MzArTrans=Request.Form("MzArTrans")
                  EnExp=Request.Form("EnExp")
                  FrExp=Request.Form("FrExp")
                  ArExp=Request.Form("ArExp")
                  MzLaExp=Request.Form("MzLaExp")
                  MzArExp=Request.Form("MzArExp")
                  EnTrans=Replace(EnTrans,"'","''")
                  FrTrans=Replace(FrTrans,"'","''")
                  ArTrans=Replace(ArTrans,"'","''")
                  MzLaTrans=Replace(MzLaTrans,"'","''")
                  MzArTrans=Replace(MzArTrans,"'","''")
                  Rs.Open "Insert Into mztable ( entrans, frtrans, artrans, mzlatrans, mzartrans, enexp, frexp, arexp, mzlaexp, mzarexp) Values (' " & EnTrans& " ',' " & FrTrans& " ',' " & ArTrans& " ',' " & MzLaTrans& " ',' " & MzArTrans& " ',' " & EnExp& " ',' " & FrExp& " ',' " & ArExp& " ',' " & MzLaExp& " ',' " & MzArExp& " ')" , Conn
                  Response.Redirect"update.asp?Ref=Main" 
                else 
                  if Ref="Search" then
                    Search=Request.Form("SearchText")                    
                    Select Case Request.Form("Lang")
                      case "En" Rs.Open "Select * from mztable WHERE entrans like '%"&Search&"%' or enexp like '%"&Search&"%'", Conn 
                      case "Fr" Rs.Open "Select * from mztable WHERE frtrans like '%"&Search&"%' or frexp like '%"&Search&"%'" , Conn 
                      case "Ar" Rs.Open "Select * from mztable WHERE artrans like '%"&Search&"%' or arexp like '%"&Search&"%'", Conn 
                      case "MzLa" Rs.Open "Select * from mztable WHERE mzlatrans like '%"&Search&"%' or mzlaexp like '%"&Search&"%'", Conn 
                      case "MzAr" Rs.Open "Select * from mztable WHERE mzartrans like '%"&Search&"%' or mzarexp like '%"&Search&"%'", Conn 
                    end Select  
                    if Rs.eof=true then 
                      Response.Redirect"update.asp?Ref=Main"
                    else%>                  
                      <html>
                      <head>
                      <SCRIPT language=javascript1.2>
        var submitOK = true;
        function SelectAll() 
  {
	len = document.SearchList.elements.length;
	var i=0;
	while(i!=len)
	{
		if (document.SearchList.elements[i].type=='checkbox') 
			document.SearchList.elements[i].checked = !document.SearchList.elements[i].checked;
		i++;
	}
  }
  function deleteMsg()
  { 
 	len = document.SearchList.elements.length;
	var i=0;
	var j=0;
	while(i!=len)
	{
		if (document.SearchList.elements[i].type=='checkbox') 
          if (document.SearchList.elements[i].checked ==true )
    	 	j++;
		i++;
	}
    if (j==0)
       alert(document.SearchList.DELETEWHAT.value);
    else   
    { 
      if(confirmDelete(document.SearchList.DELETECONFIRM.value))
      {
       	document.SearchList.DELETE.value = "1";
       	document.SearchList.submit();
      } 	
      else
	    document.SearchList.DELETE.value = "0";  
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
                      <h1 align="center">Found results</h1>
                      <form name=SearchList onsubmit="return accSub();" method=post action="update.asp?Ref=Delete">  
                      <A href="javascript:SelectAll()">All/None</A>
                      <A onclick="return (accSub())" href="javascript:deleteMsg()">Delete</A>
                      <a href="update.asp?Ref=Main">Main</a>
					  <A href="admin.asp?Ref=Log">ControlPanel</a>
                      <table border="1" width="100%" cellspacing="0">    
                      <th width="10%">N°</th>     
                      <th width="18%">En</th>     
                      <th width="18%">Fr</th>     
                      <th width="18%">Ar</th>     
                      <th width="18%">MzLa</th>     
                      <th width="18%">MzAr</th>   
                    <%do while not Rs.EOF%>  
                      <tr>
					   <td width="10%" align="left" >
                        <input type="checkbox" name="CB" value=<%=Rs("wordnu")%>>
                        <a href="update.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("wordnu")%></a>
                       </td>
                       <td width="18%" align="center"><%=Rs("entrans")%>&nbsp;</td>
                       <td width="18%" align="center"><%=Rs("frtrans")%>&nbsp;</td>
                       <td width="18%" align="center"><%=Rs("artrans")%>&nbsp;</td>
                       <td width="18%" align="center"><%=Rs("mzlatrans")%>&nbsp;</td>
                       <td width="18%" align="center"><%=Rs("mzartrans")%>&nbsp;</td>
                      </tr>
                    <%Rs.movenext
                      loop%>
                      </table>
                      <INPUT type=hidden value="Please select the word(s) you want to delete !!" name=DELETEWHAT> 
                      <INPUT type=hidden value="Are you sure you want to delete the selected word(s) ?" name=DELETECONFIRM> 
                      <INPUT type=hidden value=0 name=DELETE> 
                      <A href="javascript:SelectAll()">All/None</A>
                      <A onclick="return (accSub())"href="javascript:deleteMsg()">Delete</A>
                      <a href="update.asp?Ref=Main">Main</a>
					  <A href="admin.asp?Ref=Log">ControlPanel</a>
                      </form>
                      </body>
                      </html>
                    <%Rs.close
                    end if 'Rs=True
                  else 
                    Response.Redirect"update.asp"    
                  end if 'Ref="Search"
                end if 'Ref="New"   
              end if 'Ref="Add"     
            end if 'Ref="Modif"
          end if 'Ref=Read"   
        end if 'Ref="Delete"   
      end if 'Ref="Main"  
    end if 'ref=""   
      Set Rs = Nothing
      Conn.Close
      Set Conn = Nothing
    end if 'Enter<>"Ok"%>
	<noframes>