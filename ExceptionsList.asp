<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="ExceptionsList" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to list all the exceptions for the selected year and put them
    		 into a table.
             
-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
<%
	Dim year
	year = Request.Form("schoolYear")
%>
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

<!-- Building the Exceptions List Starting Below*************************************************************-->
<div id="menuHeader">
	<br />
		Exceptions List
</div>
<br />
<form action="" method="post" name="form1">
    <span id="instructions">
    		Click on Student Id Number to view/edit student details or to add a new student click here
        	<input type="submit" value="Add New Student" onclick="setNewAction();" style="width:250px; height:30px" />
    <br />
    	<input type="submit" value="Main Menu" onclick="setMainMenu();" style="width:250px; height:30px;"  />
    </span>
<table class="tableList" border="1">
	<th>School Year</th><th>Race</th><th>Ltr Date</th><th>Approved</th><th>SASIId</th><th>ID</th><th>Last</th><th>First</th><th>Sending School</th><th>Rec. School</th><th>Type</th>
	<%
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rs
		Set rs = connection.execute(SQL)
	%>
    <!-- Loading the table with the full race instead of an abbreviation. -->
    <%while not rs.eof%>
        <tr>
        	<td><%=rs("schoolyear")%></td>
            <!--This Condition sets the race drop down to the full race type -->
            <td><% Select Case rs("race")
					case "A"
						Response.Write("Asian")
					case "a"
						Response.Write("Asian")
					case "AI"
						Response.Write("American Indian")
					case "ai"
						Response.Write("American Indian")
					case "AP"
						Response.Write("Asian/Pacific")
					case "ap"
						Response.Write("Asian/Pacific")
					case "B"
						Response.Write("African American")
					case "b"
						Response.Write("African American")
					case "BI"
						Response.Write("Bi Racial")
					case "bi"
						Response.Write("Bi Racial")
					case "C"
						Response.Write("Caucasian")
					case "c"
						Response.Write("Caucasian")
					case "H"
						Response.Write("Hispanic")
					case "h"
						Response.Write("Hispanic")
					case "I"
						Response.Write("Indian")
					case "i"
						Response.Write("Indian")
					case "O"
						Response.Write("Other")
					case "o"
						Response.Write("Other")
				End Select	
				%></td>
            <td><%=rs("Info data entry date")%></td>
            <td><%IF rs("Approved")=True then%>Yes<%Else%>No<%End if%></td>
            <td><%=rs("SASIId")%></td>
            <td><A HREF="SdntDetail.asp?Id=<%=rs("Id")%>"</A>
            	<%=rs("Id")%></td>
            <td><%=rs("Stulast")%></td>
            <td><%=rs("StuFirst")%></td>
            <td><%=rs("Sending Sch")%></td>
            <td><%=rs("ReceivingSch")%></td>
            <td><%=rs("Type of Exception")%></td>
        <%rs.movenext%>
        </tr>
    <%wend%>
	<%
		rs.close
		Set rs = Nothing 
		connection.close
    %>
</table>  
</form>
<script>
	//This Function is used to set the action of the form.
	function setMainMenu()
	{
		document.form1.action="ExceptionsIndex.asp";
	}
	function setNewAction()
	{
		document.form1.action="AddSdntExceptions.asp";
	}
</script>
  
</body>
</html>
