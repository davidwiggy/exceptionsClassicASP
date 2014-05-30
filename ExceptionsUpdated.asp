<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to show the updated student information. 
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
<script>
	alert("Recorded Successfully updated");
</script>
<!-- Building the Main Menu*************************************************************-->
<%
	If Application("page") = "SdntDetailUpdate" Then
		setSQL()
		Dim rsUpdate
		rsUpdate = connection.execute(Application("SQL"))
		Application("page") = "SdntDetail"
	End If
%>

<%
	Dim race
		race = Request.Form("race")
		Select Case race
			Case "Asian"
				race = "A"
			Case "American Indian"
				race = "AI"
			Case "Asian/Pacific"
				race = "AP"
			Case "African American"
				race = "B"
			Case "Bi Racial"
				race = "BI"
			Case "Caucasian"
				race = "C"
			Case "Hispanic"
				race = "H"
			Case "Indian"
				race = "I"
			Case "Other"
				race = "O"
		End Select
	
	If Application("page") = "AddSdntExceptions" Then
		setSQL()
		Dim rsInsert
		Set rsInsert = connection.execute(Application("SQL"))
		Application("page") = "SdntInserted"
	End If
	
	If Application("page") = "SdntInserted" Then
		setSQL()
		Dim rsFindId
		Set rsFindId = connection.execute(Application("SQL")) '!!!!SQL QUERY TO  FIND NEW INSERT ENTRY
		Application("studentId") = rsFindId("Id")
		Application("page") = "SdntDetail" 
	End If
%>

<%
	If Application("page") = "SdntDetail" Then
		setSQL()
		Dim rs
		Set rs = connection.execute(Application("SQL"))
	End If
%>

<div class="mainMenu">
	<p id="menuHeader">
    	<br />
    	Student Information updated in the database.
    </p>

    <table class="buttonHolder"  cellspacing="15"  >
		<tr>
        	<td style="width:400px;">
        	</td>
        	<td style="width:400px;">
            	SASlld: <input type="text" id="readOnly" readonly="readonly" name="SASlld" value="<%=rs("SASIId") %>" />
            </td>
        </tr>
        <tr>
    		<td>
            	Student Last: <input type="text" id="readOnly" readonly="readonly" name="stuLast" value="<%=rs("stuLast")%>" />
            </td>
            <td>
            	Sending School: <input type="text" id="readOnly" readonly="readonly" name="sendingSch"  value="<%=rs("Sending Sch")%>"/>
            </td>
        </tr>
        <tr>
        	<td>
            	Student First: <input type="text" id="readOnly" readonly="readonly" name="stuFirst" value="<%=rs("stuFirst")%> " />
            </td>
            <td>
            	Receiving School: <input type="text" id="readOnly" readonly="readonly" name="recSch" value="<%=rs("ReceivingSch")%>"  />
        </tr>
        <tr>
        	<td>
            	Parent Last: <input type="text" id="readOnly" readonly="readonly" name="parLast" value="<%=rs("Parlast")%>"/>
            </td>
            <td>
            	Grade: <input type="text" id="readOnly" readonly="readonly" name="grade" value="<%=rs("Grade")%>"  />
            </td>
        </tr>
        <tr>
        	<td>
            	Parent First: <input type="text" id="readOnly" readonly="readonly" name="parFirst" value="<%=rs("Parfirst")%>"/>
            </td>
            <td>
            	Race: <input type="text" id="readOnly" readonly="readonly" name="race" value="<%=rs("race")%>" style="max-width:25px;" />
        </tr>
      	<tr>
        	<td>
            	Address: <input type="text" id="readOnly" readonly="readonly" name="address" value="<%=rs("street")%>" style="min-width:200px;" />
            </td>
            <td>
            	<!--This is setting up the radio inputs with the current schools approved status. -->
            	Approved: <input type="radio" id="readOnly" name="approved" value="Yes"<% If rs("Approved") = True then
																					%>checked />Yes<%
																			 	  Else
																			 		%>disabled="disabled"/>Yes<%
																	   		 End If%>
                                <input type="radio" id="readOnly" name="approved" value="No" <% If rs("Approved") = False then
																				  	%>checked />No<%
																				  Else
																				  	%>disabled="disabled"/>No<%	
																			End If%>
            </td>
        </tr>
        <tr>
        	<td>
            	City: <input type="text" id="readOnly" name="city" value="<%=rs("City")%>" style="min-width:200px;" readonly="readonly" />
            </td>
            <td>
            	Renewal: <input type="radio" id="readOnly" name="renewal" value="Yes" disabled="disabled" />Yes
                		 <input type="radio" id="readOnly" name="renewal" value="No" checked  />No
            </td>
        </tr>
        <tr>
        	<td>
            	State: <input type="text" id="readOnly" name="state" value="SC" readonly="readonly" style="max-width:60px;" />
            </td>
            <td>
            	Zip: <input type="text" id="readOnly" name="zip" value="<%=rs("Zip")%>" style="max-width:60px;" readonly="readonly"/>
            </td>
        </tr>
        <tr>
        	<td>
            	Exception Type: <input type="text" id="readOnly" name="exceptionType" value="<%=rs("Type of Exception") %>" readonly="readonly" b/>
            </td>
        </tr>
        <tr>
        	<td>
            	Data Entry Date: <input type="text" id="readOnly" name="entryDate" value="<%=rs("Info data entry date")%>" style="max-width:85px;" readonly="readonly" />
            </td>
            <td>
            	Year: <input type="text" name="schYear" value="<%=rs("SchoolYear")%>" id="readOnly" readonly="readonly" />
            </td>
        </tr>
	</table>
    <br />
    <form action="" method="post" name="formExceptionsUpdated" >
        <div id="updateButton">
        	<input type="hidden" value="<%=rs("SchoolYear")%>" name="schoolYear"  />
        	<input type="submit" value="Return to Student Update" onclick="setExceptionsAction();" style=" width:300px; height:40px;" />
            <input type="submit" value="Return to Exceptions List" onclick="setExceptionsListAction();" style="width:300px; height:40px;"/>
            <input type="submit" value="Return to Main Menu" onclick="mainMenu();"  style="width:300px; height:40px;"/>
        </div>
	</form>
</div>
<script>
	//These function set the form action depending on the button click
	function setExceptionsAction()
	{                                
		document.formExceptionsUpdated.action="SdntDetail.asp?Id=<%=rs("Id")%>";
	}
	function setExceptionsListAction()
	{
		document.formExceptionsUpdated.action="ExceptionsList.asp";
	}
	function mainMenu()
	{
		document.formExceptionsUpdated.action="ExceptionsIndex.asp";
	}
</script>
</body>
</HTML>