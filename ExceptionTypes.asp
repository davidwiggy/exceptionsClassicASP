<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="ExceptionTypes" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to preview the current types of exceptions.
             
-->

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
	'This is calling a function from the .inc file to set the sql. It also executes the sql and puts the results into
	'a result set(rs)
	setSQL()
	Dim SQL
	SQL = Application("SQL")
	Dim rs
	Set rs = connection.execute(SQL)
%>

<!-- Building the Main Menu*************************************************************-->
<div class="mainMenu">
	<p id="menuHeader">
    	Exceptions Types
    </p>
<form action="ExceptionsIndex.asp" method="post" name="form1">
    <table class="buttonHolder" border="1" >

            <%while not rs.eof%>
            <tr>
         		<td>
     			<% 
					Response.Write(rs("Type"))
	 			%>
        		</td>
    
        	
        <%
			rs.movenext
		%>
    <%wend%>
    </tr>
	<%
		rs.close
		Set rs = Nothing 
		connection.close
    %>
	</table>
    <p id="updateButton">
   <input type="submit" value="Main Menu" style="width:300px; height:40px;"> 
   </p>
</form>
</div>
</body>
</HTML>