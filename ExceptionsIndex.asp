<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="ExceptionsIndex"  
   Application("studentId") = ""
   Application("labelDate") = ""
   Application("SQL") = ""
%>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this program is to allow the user to access the exceptions database. It allows the user
    		 to print denied and approved letters and labels. It allows the user to add to the exceptions database. It also
             allows the user to edit current students in the database.
             
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

<!-- Building the Main Menu*************************************************************-->
<div class="mainMenu">
	<p id="menuHeader">
    	Attendance Exceptions
    </p>
<form action="" method="post" name="form1">
    <table class="buttonHolder">
		<tr>
            <td id="schoolYear">
            	School Year:  <select name="schoolYear">
                <%
					getYears()
				%></select>
            </td>
   		 </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<br />
            		<input type="submit" value="Exceptions" onclick="setExceptionsList();" >
              
            </td>
            
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<input type="submit" value="Letters and Labels" onclick="setLettersAndLabels();">
            </td>
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<br />
                <br />
            	<input type="submit" value="Exception Types" onclick="setExceptionTypes();">
            </td>
            
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<input type="submit" value="Exceptions By Selected Receiving School" onclick="setReceivingSch();">
            </td>
         </tr>
	</table>
      </form>
</div>
<script>
	//These function set the actions on the form depending on which button is clicked
	function setExceptionsList()
	{
		document.form1.action="ExceptionsList.asp";
	}
	function setLettersAndLabels()
	{
		document.form1.action="LettersAndLabels.asp";
	}
	function setExceptionTypes()
	{
		document.form1.action="ExceptionTypes.asp";
	}
	function setReceivingSch()
	{
		document.form1.action="ReceivingSchool.asp";
	}
</script>
<%
	'This function loads a drop down list with years.
	sub getYears()
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rs
		Set rs = connection.execute(SQL)

		While Not rs.EOF
		%><option value="<%=rs("SchoolYear")%>"><%=rs("SchoolYear")%></option>
		<%
		rs.MoveNext
		Wend
		
		rs.close
		Set rs = Nothing 
		connection.close
	end sub
%>

</body>
</HTML>