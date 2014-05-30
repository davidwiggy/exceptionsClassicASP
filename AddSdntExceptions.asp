<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="AddSdntExceptions" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to add students to the exceptions database.
             
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
    	Add Student to Exceptions List
    </p>
<form action="" method="post" name="formSdntDetail" >
    <table class="buttonHolder"  cellspacing="15"  >
		<tr>
        	<td style="width:400px;">
           		<span style="font-weight:bold;"> Yellow Fields are Required. </span>
        	</td>
        	<td style="width:400px;">
            	SASIId: <input type="text" name="sas" value="" />
            </td>
        </tr>
        <tr>
    		<td>
            	Student Last: <input type="text" name="stuLast" value="" style="background-color:#FFFF00;" />
            </td>
            <td>
            	Sending School: <input type="text" name="sendingSch"  value=""/>
            </td>
        </tr>
        <tr>
        	<td>
            	Student First: <input type="text" name="stuFirst" value="" style="background-color:#FFFF00;"  />
            </td>
            <td>
            	Receiving School: <select name="receivingSch">
                <%
					getRecSch()
					Application("page") = "AddSdntExceptions"
				%></select>
        </tr>
        <tr>
        	<td>
            	Parent Last: <input type="text" name="parLast" value="" style="background-color:#FFFF00;"/>
            </td>
            <td>
            	<!--This is building an array with all the grade levels. Then it uses this array to load a drop down list. -->
            	Grade: <select name="grade">
                <%
					gradeArray = Array("K-4", "Kindergarten", "1st", "2nd", "3rd", "4th", "5th", "6th", "7th", "8th", "9th", "10th", "11th", "12th")
					For x = 0 To 13
						%><option value="<%Response.Write(gradeArray(x))%>"><%response.Write(gradeArray(x))%></option><%
					Next
					%></select>
            </td>
        </tr>
        <tr>
        	<td>
            	Parent First: <input type="text" name="parFirst" value="" style="background-color:#FFFF00;"/>
            </td>
            <td>
            	 Race: <select name="race">
                 <%
				 	raceArray = Array("Asian", "American Indian", "Asian/Pacific", "African American", "Bi Racial", "Caucasian", "Hispanic", "Indian", "Other")
					For x = 0 To 8
						%><option value="<%Response.Write(raceArray(x))%>"><%Response.Write(raceArray(x))%></option><%
					Next
					%></select>
            </td>
        </tr>
      	<tr>
        	<td>
            	Address: <input type="text" name="address" value="" style="min-width:200px; background-color:#FFFF00;" />
            </td>
            <td>
				Zip: <input type="text" name="zip" value="" style="max-width:60px;background-color:#FFFF00;" />
            </td>
        </tr>
        <tr>
        	<td>
            	City: <input type="text" name="city" value="" style="min-width:200px;background-color:#FFFF00;" />
            </td>
            <td>
            	<!--This is setting up the radio inputs with the current schools approved status. -->
            	Approved: <input type="radio" name="approved" value="Yes" />Yes
						  <input type="radio" name="approved" value="No" checked="checked" />No
            </td>
        </tr>
        <tr>
        	<td>
            	State: <input type="text" id="readOnly" name="state" value="SC" readonly="readonly" style="max-width:60px;" />
            </td>
            <td>
            	Renewal: <input type="radio" name="renewal" value="Yes" />Yes
                		 <input type="radio" name="renewal" value="No" checked  />No
            </td>
        </tr>
        <tr>
        	<td>
            	Exception Type: <select name="exception">
                <%
					setExceptions()
					application("page") = "AddSdntExceptions"
				%></select>
            </td>
        </tr>
        <tr>
        	<td>
            	Data Entry Date(Format: mm/dd/yyyy): <input type="text" name="entryDate" value="" maxlength="10" style="max-width:85px;" />
            </td>
            <td>
            	Year: <select name="year">
                <%
					setYears()
					application("page") = "AddSdntExceptions"
				%></select>
            </td>
        </tr>
	</table>
    <br />
    <div id="updateButton">
    	<div id="buttonCells">
    		<input type="submit" value="Insert Student Information" onclick=" return checkFormInfo();" style="width:300px; height:40px;" />
            <input type="submit" value="Main Menu" onclick="mainMenu();" style="width:300px; height:40px;"  />
        </div>
    </div>
</form>
</div>

<!--This function updates the current students information in the database from the information in the 
	form. It ONLY VALIDATES student first and last name. It validates this information using javascrip
    and if the boxes contains values it updates the db, if they do not it alerts the user with prompt
    windows.  -->
<script>
	//This function sets the action of the form.
	function mainMenu()
	{
		document.formSdntDetail.action="ExceptionsIndex.asp";
	}
	
	//This function validates the information entered into the form.
	function checkFormInfo()
	{
			var first    =document.forms["formSdntDetail"]["stuFirst"].value;
			var last     =document.forms["formSdntDetail"]["stuLast"].value;
			var city     =document.forms["formSdntDetail"]["city"].value;
			var zip      =document.forms["formSdntDetail"]["zip"].value;
			var parFirst =document.forms["formSdntDetail"]["parFirst"].value;
			var parLast  =document.forms["formSdntDetail"]["parLast"].value;
			var address  =document.forms["formSdntDetail"]["address"].value;
			var entry    =document.forms["formSdntDetail"]["entryDate"].value;
			var pattern =/^([0-9]{2})\/([0-9]{2})\/([0-9]{4})$/;

			
			if(first===null || first==="" || last===null || last==="")
			{
				alert("Both First and Last Name Fields Must be Filled out!");
				return false;
			}
			else if(parFirst===null || parFirst==="" || parLast===null || parLast==="")
			{
				alert("Both Parent First and Last Name Fields Must be Filled out!");
				return false;
			}
			if(address===null || address==="")
			{
				alert("Address Field Must be Filled out!");
				return false;
			}
			else if(city===null || city==="" || /^\D+$/.test(city) == false)
			{
				alert("City must be filled out and not contain any numbers.");
				return false;
			}
			else if(zip===null || zip==="" || /^\d+$/.test(zip) == false || zip.length != 5)
			{
				alert("Zip code must be filled in, contain only numbers, and have be 5 digits long.")
				return false;
			}
			else if(entry===null || entry==="" || pattern.test(entry) == false)
			{
				alert("Data Entry Date Must be Entered in this Format MM/DD/YYYY. \n\nCHECK YOUR MONTH MUST BE (MM) TWO DIGITS \n\n EXAMPLE: 01/01/2001")
				return false;
			}
			else if(isValidDate(entry) == false)
			{
				alert("Date must be a valid date");
				return false;
			}
			else
			{	
				<% application("page") = "AddSdntExceptions" %>
				document.formSdntDetail.action="ExceptionsUpdated.asp";
				return true;
			}
	}
	
	//This function validates the date
	function isValidDate(date) {
        var valid = true;

        //date = date.replace('/-/g', '');

        var month = parseInt(date.substring(0, 2));
        var day   = parseInt(date.substring(3, 5));
        var year  = parseInt(date.substring(6, 10));
		
        if((month < 1) || (month > 12)) valid = false;
        else if((day < 1) || (day > 31)) valid = false;
        else if(((month == 4) || (month == 6) || (month == 9) || (month == 11)) && (day > 30)) valid = false;
        else if((month == 2) && (((year % 400) == 0) || ((year % 4) == 0)) && ((year % 100) != 0) && (day > 29)) valid = false;
        else if((month == 2) && ((year % 100) == 0) && (day > 29)) valid = false;

    	return valid;
	}
	
	
</script>    
<!--This is Setting the exceptions drop down box with all the types of exceptions. . -->
<%
	sub setExceptions()
		Application("page") = "SdntDetail3"
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rsExceptions
		Set rsExceptions = connection.execute(SQL)
		
		While Not rsExceptions.EOF
				%><option value="<%Response.Write(rsExceptions("Type"))%>"><%Response.Write(rsExceptions("Type"))%></option><%
			rsExceptions.MoveNext
		Wend
		rsExceptions.close
		Set rsExceptions = Nothing
	End sub
%>

<!--This is setting the receiving school drop down box with all the schools from the district. -->
<%
	sub getRecSch()
		Application("page") = "SdntDetail2"
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rsSchool
		Set rsSchool = connection.execute(SQL)
		
		While Not rsSchool.EOF
				%><option value="<%Response.Write(rsSchool("SchoolInit"))%>"><%Response.Write(rsSchool("SchoolInit"))%></option><%
			rsSchool.MoveNext
		Wend
		rsSchool.close
		Set rsSchool = Nothing 
	end sub
%>
<!--This is setting the years in the drop down list -->
<%
	sub setYears()
		Application("page") = "SdntDetail4"
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rsYears
		Set rsYears = connection.execute(SQL)
		
		While Not rsYears.EOF
			%><option value="<% Response.Write(rsYears("SchoolYear"))%>"><% Response.Write(rsYears("SchoolYear")) %></option><%
		rsYears.MoveNext
		Wend
		
		rsYears.close
		Set rsYears = Nothing 
		connection.close
	end sub
%>

</body>
</html>
