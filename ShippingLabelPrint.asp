<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
   <link rel="stylesheet" type="text/css" href="ExceptionsCSS.css" />
   <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
   <meta name="GENERATOR" content="Microsoft FrontPage 4.0">
   
	<title>BCSD Intranet Home Page</title>
<!-- STANDARD HEADER BEGIN******************************************************************-->
</head>
<body>
<table width="100%">
<tr><td align="right" valign="top"><b><A HREF="/index.asp"><font face="tahoma" color="#006600" size="0">Home Page</font></b></a></td></tr>
<br>

</table>

<table width="100%">
<tr>
  	<td align="left"> <img src="BCSD_Logo_Medium.jpg" alt="Logo" width="143" height="120" /></td>
	<td align="center" valign="middle"><b><font face="tahoma" color="#006600" size="6">Welcome to Berkeley County School District's Intranet</font></b></td>
</tr>
</table>
<hr>

<div align="center">
<font face="tahoma" color="#006600" size="1"><b>
<% =formatDateTime(date(), vblongdate) %></b></font><br>

</div>
<%
	If Application("page") <> "ShippingLabelPrint" Then
		session("shipName") = Request.Form("shippingName")
		response.Write(session("shipName"))
		Application("page") = "ShippingLabelPrint"
	End If
%>
<%
	setSQL()
	Dim SQL
	SQL = Application("SQL")
	Dim rs
	Set rs = connection.execute(SQL)
%>

<!-- Building the Main Menu*************************************************************-->
<div class="mainMenu">
<form action="" method="post" name="form1" >
	<p id="menuHeader" >
    	Shipping Label
    </p>
    <table class="labelLayout" bgcolor="#FFFFFF" cellspacing="35px" id="printTable" border="1">
     	<% 
	 		names = rs("Parfirst") & " " & rs("Parlast")
			city = rs("City") & ", " & rs("State") & " " & rs("Zip")
	 	%>
        <tr>
        	<td><%=names%><br /><%=rs("Street")%><br /><%=city%></td>   
      	</tr>  
	<%
		rs.close
		Set rs = Nothing 
		connection.close
    %>
	</table>
   
    <input type="submit" value="Print" onclick="printData()" />
      </form>
</div>

<script>
	function printData()
	{
	   var divToPrint=document.getElementById("printTable");
	   newWin= window.open("");
	   newWin.document.write(divToPrint.outerHTML);
	   newWin.print();
	   newWin.close();
	}
</script>


</body>
</HTML>