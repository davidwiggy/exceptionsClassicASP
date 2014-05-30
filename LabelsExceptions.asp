<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LabelsExceptions" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to preview the labels for the selected date. 
    		 It also provides a button that access the print labels file.
             
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
	Application("labelDate") = Request.Form("LetterDates")
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
    	Labels Exceptions
    </p>
    <p id="selectionHeaderLabels">
    	The current selected date is: <%=(Application("labelDate"))%><br />
    </p>
      <div id="updateButton">
    	<input type="submit" value="Main Menu" onclick="setMain();" style="width:250px; height:35px;"  />
        <input type="submit" value="Letters and Labels" onclick="setLetters();" style="width:250px; height:35px;"  />
    	<input type="submit" value="Print Labels" onclick="setPrint();" style="width:250px; height:35px;"  />
    </div>
   
    <p id="selectionHeaderLabels">
    </p>
    <table class="labelsTable"  border="1">
		 <%while not rs.eof%>
     	<% 
				 If Not rs.EOF Then
					If rs("Zip")=null Or rs("Zip")="" Or rs("Zip")=0 Then
						zipTemp = ""
					else
						zipTemp = rs("Zip")
					End If
					names = rs("Parfirst") & " " & rs("Parlast")
					city = rs("City") & ", " & rs("State") & " " & zipTemp
	 	%>
        <tr>
        	<td ><%=names%><br /><%=rs("Street")%><br /><%=city%></td>   
		<% End If %>
        <%If Not rs.EOF Then
			rs.movenext
		  End If	
		%>
     	<% 
				 If Not rs.EOF Then
					If rs("Zip")=null Or rs("Zip")="" Or rs("Zip")=0 Then
						zipTemp = ""
					else
						zipTemp = rs("Zip")
					End If
					names = rs("Parfirst") & " " & rs("Parlast")
					city = rs("City") & ", " & rs("State") & " " & zipTemp
	 	%>
           	<td ><%=names%><br /><%=rs("Street")%><br /><%=city%></td>
        <%  
			End If
		%>
        <%If Not rs.EOF Then 
			rs.movenext 
			End If
		%>
     	<% 
				 If Not rs.EOF Then
					If rs("Zip")=null Or rs("Zip")="" Or rs("Zip")=0 Then
						zipTemp = ""
					else
						zipTemp = rs("Zip")
					End If
					names = rs("Parfirst") & " " & rs("Parlast")
					city = rs("City") & ", " & rs("State") & " " & zipTemp
	 	%>
            <td ><%=names%><br /><%=rs("Street")%><br /><%=city%></td>
        <%  
			End If
			If Not rs.EOF Then
				rs.movenext
			End If
		%>
      	</tr>  
    </div>
    <%wend%>

	<%
		rs.close
		Set rs = Nothing 
		connection.close
    %>
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
		document.form1.action="LabelsPrint.asp"
	}
	
</script>
</body>
</HTML>