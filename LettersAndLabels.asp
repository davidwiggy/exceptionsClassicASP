<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LettersAndLabels" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to provide the user with a menu regarding previewing and
             printing labels, and letters.
             
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
    	Letters And Labels
    </p>
<form action="" method="post" name="form1">
    <table class="buttonHolder">
		<tr>
            <td id="schoolYear">
            	Select Letter Date to Print Letters:  <select name="letterDates">
                <%
					getYears()
				%></select>
            </td>
   		 </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<br />
            		<input type="submit" value="Preview Approved and Print" onclick="setPreApproved();" />
              
            </td>
            
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<input type="submit" value="Preview Denied and Print" onclick="setPreDenied();"/>
            </td>
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align: middle">
            	<input type="submit" value="Preview Labels and Print" onclick="setPreLabels();"  />
            </td>     
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align:middle">
            	<input type="submit" value="Print All: Labels, Denied Letters, and Approved Letters" onclick="setPrintAll();"  />
            </td>
         </tr>
         <tr id="buttonCells" align="center">
         	<td style="vertical-align:middle">
            	<input type="submit" value="Main Menu" onclick="mainMenu();" />
            </td>
         </tr>
	</table>
      </form>
</div>
<script>
	//This functions set the action of the form depending on which button is clicked
	function setPrintAll()
	{
		if(checkSelection())
		{
			alert("For Labels use Avery Easy Peel Labels (Template 5160).");
			window.open("LabelsPrint.asp");
			window.open("LetterApprovedPrint.asp");
			window.open("LetterDeniedPrint.asp");
		}
	}
	
	function mainMenu()
	{
		document.form1.action="ExceptionsIndex.asp";
	}
	function setPreApproved()
	{
		var flag = checkSelection();
		if(flag == true)
		{
			document.form1.action="LetterApproved.asp";
		}
		
	}
	
	function setPreDenied()
	{
		var flag = checkSelection();
		if(flag == true)
		{
			document.form1.action="LetterDenied.asp";
		}
	}
	
	function setPreLabels()
	{
		var flag = checkSelection();
		if(flag == true)
		{
			document.form1.action="LabelsExceptions.asp";
		}
	}
	
	function printLabels()
	{
		var flag = checkSelection();
		if(flag == true)
		{
			alert("Use Avery Easy Peel Labels (Template 5160).");
			document.form1.action="LabelsPrint.asp";
		}
	}
	
	function printApprovedLetters()
	{
		var flag = checkSelection();
		if(flag == true)
		{
			document.form1.action="LetterApprovedPrint.asp"
		}
	}
	
	function printDeniedLetters()
	{
		if(checkSelection())
		{
			document.form1.action="LetterDeniedPrint.asp"
		}
	}
	
	//This function checks to make sure that the user has selected a date before preceding. 
	function checkSelection()
	{
		var date=document.forms["form1"]["letterDates"].value;
				
			if(date===null || date==="")
			{
				alert("You must select a date first");
				document.form1.action="";
				return false;
			}
			else
			{	
				return true;
			}
	}
</script>
<%
	'This function loads a drop down list from a record set
	sub getYears()
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rs
		Set rs = connection.execute(SQL)

		While Not rs.EOF
		%><option value="<%=rs("Info data entry date")%>"><%=rs("Info data entry date")%></option>
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