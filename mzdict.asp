<%Set Conn = Server.CreateObject("ADODB.Connection")      
Conn.Open "DBQ=" & Server.MapPath("db/mzdict.mdb") & ";Driver={Microsoft Access Driver (*.mdb)};Driverld=25","","azerty"
Set Rs = Server.CreateObject("ADODB.Recordset")
Ref=Request.Querystring("Ref")
    if Ref="" then
        Rs.Open "Select * From mztable" , Conn%>
        <html>
        <body  background="images/berriane.jpg">
		<!--#include file="menu.asp"-->
		<link rel="stylesheet" href="css/style.css" type="text/css">
		<br><h1 align="center">The Mozabite Dictionary</h1>
        <form method=post action="mzdict.asp?Ref=Search">
        <p align="center" dir="ltr">
        Language :
        <select size="1" name="Lang">
          <option selected>En</option>
          <option>Fr</option>
          <option>Ar</option>
          <option>MzLa</option>
          <option>MzAr</option>
        </select>
        Search :                          
        <input type="text" name="SearchText" size="20"> 
        <input type="submit" value="Search"></p>
        <p>Select the language and type your word then click the search button to see the result</p>
                  <p>
          En: English<br>
          Fr: French<br>
          Ar: Arabic<br>
          Mz (La): Mozabite written with latin letters<br>
          Mz (Ar): Mozabite written with arabic letters
          </p>
        </form>
      <%else
          if Ref="Read" then
            WordRef=Request.Querystring("reference")
            WordRef=CINT(WordRef)
            Rs.Open "Select * from mztable WHERE wordnu=" & WordRef , Conn%>
          <html>
          <head><title>Mozabite Dictionary !!</title></head>
          <body  background="images/berriane.jpg">
		  <!--#include file="menu.asp"-->
		  <link rel="stylesheet" href="css/style.css" type="text/css">
          <br><p class="lien" align="center"><a href="mzdict.asp">Search</a>
		  <a href="main.asp">Main</a></p>
          <table width="40%" align="center">
          <th colspan="2">Translations</th>
          <tr>
          <td><b>English:</b></td>
          <td><%=Rs("EnTrans")%></td>
          </tr>
          <tr>
          <td><b>French:</b></td>
          <td><%=Rs("FrTrans")%></td>
          </tr>
          <tr>
          <td><b>Arabic:</b></td>
          <td><%=Rs("ArTrans")%></td>
          </tr>
          <tr>
          <td><b>Mozabite (Lt):</b></td>
          <td><%=Rs("MzLaTrans")%></td>
          </tr>
          <tr>
          <td><b>Mozabite (Ar):</b></td>
          <td><%=Rs("MzArTrans")%></td>
          </tr>
          </table><br>
          <table width="80%" align="center">
          <th colspan="2">Explications</th>
          <tr>
          <td><b>English: </b> </td>
          <td><%=Rs("EnExp")%></td>
          </tr>
          <tr>
          <td><b>French:</b></td>
          <td><%=Rs("FrExp")%></td>
          </tr>
          <tr>
          <td><b>Arabic:</b></td>
          <td><%=Rs("ArExp")%></td>
          </tr>
          <tr>
          <td><b>Mozabite (La):</b></td>
          <td><%=Rs("MzLaExp")%></td>
          </tr>
          <tr>
          <td><b>Mozabite (Ar):</b></td>
          <td><%=Rs("MzArExp")%></td>
          </tr>
          </table>
          <p>
          En: English<br>
          Fr: French<br>
          Ar: Arabic<br>
          Mz (La): Mozabite written with latin letters<br>
          Mz (Ar): Mozabite written with arabic letters
          </p>
          <p class="lien" align="center"><a href="mzdict.asp">Search</a>
		   <a href="main.asp">Main</a></p>
          </body>
          </html>
          <%Rs.Close
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
                    if Rs.eof=true then%> 
                      <html>
                      <body  background="images/berriane.jpg">
					  <!--#include file="menu.asp"-->
					  <link rel="stylesheet" href="css/style.css" type="text/css">
					  <br><h1 align="center">No results found</h1>   
					  <p class="lien" align="center"><a href="mzdict.asp">Try agin !!</a>
					  <a href="main.asp">Main</a></p>
					<%else%>                 
                      <html>
                      <body  background="images/berriane.jpg">
					  <!--#include file="menu.asp"-->
					  <link rel="stylesheet" href="css/style.css" type="text/css">
                      <br><h1 align="center">Found results</h1>                
                      <table border="1" width="100%" cellspacing="0">    
                      <th width="20%">En</th>     
                      <th width="20%">Fr</th>     
                      <th width="20%">Ar</th>     
                      <th width="20%">Mz (La)</th>     
                      <th width="20%">Mz (Ar)</th>   
                    <%do while not Rs.EOF%>  
                      <tr>
                       <td width="20%" align="center"><a href="mzdict.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("entrans")%></a>&nbsp;</td>
                       <td width="20%" align="center"><a href="mzdict.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("frtrans")%></a>&nbsp;</td>
                       <td width="20%" align="center"><a href="mzdict.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("artrans")%></a>&nbsp;</td>
                       <td width="20%" align="center"><a href="mzdict.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("mzlatrans")%></a>&nbsp;</td>
                       <td width="20%" align="center"><a href="mzdict.asp?Ref=Read&reference=<%=Rs("wordnu")%>"><%=Rs("mzartrans")%></a>&nbsp;</td>
                      </tr>
                    <%Rs.movenext
                      loop%>
                      </table>
                      <p class="lien" align="center"><a href="mzdict.asp">Search</a>
					  <a href="main.asp">Main</a></p>
                      </body>
                      </html>
                    <%Rs.close
                    end if 'Rs=True  
                  end if 'Ref="Search"
                end if 'Ref="New"   
              end if 'Ref="Add"     
      Set Rs = Nothing
      Conn.Close
      Set Conn = Nothing%>
    <noframes>
