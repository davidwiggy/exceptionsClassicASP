<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LetterApproved" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to preview one approved letters for the selected date. 
    		 It also provides a button that access the print approved labels file.
             
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
<!--Setting the global variable for labelDate-->
<%
	Application("labelDate") = Request.Form("LetterDates")
%>

<!--Calling the setSQL function in the .inc file, that sets the sql. Then the sql is executed and the
	results is put into the result set(rs)
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
    	Approved Letter Preview
    </p>
    <div id="updateButton">
    	<input type="submit" value="Main Menu" onclick="setMain();" style="width:250px; height:35px;"  />
        <input type="submit" value="Letters and Labels" onclick="setLetters();" style="width:250px; height:35px;"  />
    	<input type="submit" value="Print Approved" onclick="setPrint();" style="width:250px; height:35px;"  />
    </div>
    <table class="previewTable" bgcolor="#FFFFFF" >
        <tr>
        	<td><%=rs("Info data entry date")%></td>
      	</tr>  
        <tr>
        	<td><br /><%=rs("Parfirst")%><%=(" ")%><%=rs("Parlast")%></td>
        </tr>
        <tr>
        	<td><%=rs("Street")%></td>
        </tr>
        <tr>
        	<td><%=rs("City")%><%If rs("City") <> "" Then Response.Write(", ") End If %><%=rs("State")%><%=(" ")%><%if rs("Zip") <> 0 Then 
																			                                           Response.Write(rs("Zip")) 
																	                                                End If%></td>
        </tr>
        <tr>
        	<td><br /><%=rs("Stufirst")%><%=(" ")%><%=rs("Stulast")%></td>
        </tr>
        <tr>
        	<td><%=rs("SchoolName")%></td>
        </tr>
        <tr>
        	<td><%=("Approval for Exceptions: ")%><%=rs("Type of exception")%>
        </tr>
        <tr>
        	<td><p>Dear Sir/Madame:<br /><br />

				Your request for an exception to the above mentioned child’s/children’s regular geographic school assignment has been given careful administrative review.<br /><br />

				After thorough consideration of all evidence and documentation presented in support of this petition, your request has been APPROVED as being in compliance with the recognized categories of exceptions 		  authorized by Berkeley County School Board policies.<br /><br />

		This approval is only for the <%=rs("Schoolyear")%> school year.  Request for exceptions must be made on an annual basis through my office.<br />
        <br />If we can answer any further questions or assist you in any way, please do not hesitate to call.<br />
 		<br />Sincerely,
        <br /><br />Charlie Davis, Administrative Assistant for Superintendent
        <br /><br />Division of Administration and Pupil Services
        <br /><br />BD/ps</p></td>
        </tr>
        <tr>
        	<td><br />C: <%=rs("SchoolPrincipal")%></td>
        </tr>
        <tr>
        	<td>File</td>
        </tr>
    </div>

	
	</table>
</form>

<script>
	//These are functions that determine the action of the form based on the button that is clicked
	function setMain()
	{
		document.form1.action="ExceptionsIndex.asp"
	}
	
	function setLetters()
	{
		document.form1.action="LettersAndLabels.asp"
	}
	
	function setPrint()
	{
		document.form1.action="LetterApprovedPrint.asp"
	}
</script>
</body>
</HTML>