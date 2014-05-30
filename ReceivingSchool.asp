<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="ReceivingSchool" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to view all the exceptions by school per year. 
    		 It is loaded dynamically depending on the user selection.
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
	flag = false
	If Application("resSchLoaded") = "loaded" Then
		Application("page") = "ReceivingSchool2"
		school = Request.Form("recSch")
		schYear = Request.Form("year")
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rs
		Set rs = connection.execute(SQL)
		Application("resSchLoaded") = "notLoaded"
		Application("page") = "ReceivingSchool"
		flag = true
	End If
%>

<!-- Building the Main Menu*************************************************************-->
<div class="mainMenu">
	<p id="menuHeader">
    	Exceptions Listing - Select Year and Receiving School
    </p>
<form action="" method="post" name="form1">
    <table class="buttonHolder" >
    	<tr>
        	<td id="updateButton">Receiving School: <select name="recSch">
                <%
					setSchools()
				%></select></td>
            <td id="updateButton">School Year: <select name="year">
            	<%
					getYears()
				%></select></td>
        </tr>
        <tr>
        	<td id="updateButton"><input type="submit" value="Submit" onclick="return validation();" style="width:250px; height:35px;" />
        	<td id="updateButton"><input type="submit" value="Main Menu" onclick="setAction();" style="width:250px; height:35px;" />
        </tr>
    </table>
    <table class="buttonHolder"  border="1">
    	<!-- This flag is used when the page is loaded. If the flag is false the table is not loaded if it is true the table is loaded.-->
		<% If flag = true Then %>
        <th align="left">School Year</th><th align="left">Student Last</th><th align="left">Student First</th><th align="left">Grade</th><th align="left">Race</th><th align="left">Sending School</th><th align="left">Approved</th><th align="left">Type</th>
            <%while not rs.eof%>
            <tr>
				<td>
                <%
					=rs("SchoolYear")
				%>
                </td>
         		<td>
     			<% 
					=(rs("Stulast"))
	 			%>
        		</td>
    			<td>
				<%
					=(rs("Stufirst"))
				%>
                </td>
          		<td>
				<%
					=(rs("Grade"))
				%>
                </td>
                <td>
				<%
					=(rs("Race"))
				%>
                </td>
                <td>
				<%
					=(rs("Sending Sch"))
				%>
                </td>
      			<td>
				<%
					'This is test the approved rs and the printing out yes/no depending on the outcome of the test.
					If rs("Approved") = True Then 
						Response.Write("Yes") 
					Else 
						Response.Write("No") 
					End If
				%>
                </td>
       			<td>
				<%
					=(rs("Type of Exception"))
				%>
                </td>
        	
        <%
			rs.movenext
		%>
    <%wend%>
    <% End If %>
    </tr>
	<%
		If flag = true Then
			rs.close
			Set rs = Nothing 
			connection.close
		End if
    %>
	</table>
    <p id="updateButton">
   
   </p>
</form>
</div>
<script>
	//This function is setting the action of the form
	function setAction()
	{
		document.form1.action="ExceptionsIndex.asp";
	}
	
	//This function is checking whether the user has made selections
	function validation()
	{
		var schoolYear=document.forms["form1"]["year"].value;
		var school = document.forms["form1"]["recSch"].value;
				
			if(schoolYear===null || schoolYear==="" || school===null || school==="")
			{
				alert("You must select school and year");
				<% 
					Application("resSchLoaded") = "notLoaded"
				%>
				return false;
			}
			else
			{
				<%
					Application("resSchLoaded") = "loaded"
					Application("page") = "ReceivingSchool2"
				%>
				return true;
			}
	}


</script>
<%
	'This function is loading a drop down list with the current schools.
	sub setSchools()
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rsSch
		Set rsSch = connection.execute(SQL)

		While Not rsSch.EOF
		%><option value="<%=rsSch("ReceivingSch")%>"><%=rsSch("ReceivingSch")%></option>
		<%
		rsSch.MoveNext
		Wend
		
		rsSch.close
		Set rsSch = Nothing 
	end sub

	'This function is loading a drop down list with years.
	sub getYears()
		Application("page") = "ExceptionsIndex"
		setSQL()
		Dim SQL
		SQL = Application("SQL")
		Dim rsYears
		Set rsYears = connection.execute(SQL)
		%><option value=""></option><%
		
		While Not rsYears.EOF
		%><option value="<%=rsYears("SchoolYear")%>"><%=rsYears("SchoolYear")%></option>
		<%
		rsYears.MoveNext
		Wend
		
		rsYears.close
		Set rsYears = Nothing 
		Application("page") = "ReceivingSchool"
	end sub
%>
</body>
</HTML>