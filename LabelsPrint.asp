<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LabelsPrint" %>
<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to print the labels for the
    		 selected date.
             
-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>

<meta http-equiv=Content-Type content="text/html; charset=us-ascii">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-536870145 1107305727 0 0 415 0;}
@font-face
	{font-family:"Segoe UI";
	panose-1:2 11 5 2 4 2 4 2 2 3;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-520084737 -1073683329 41 0 479 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;}
p.MsoAcetate, li.MsoAcetate, div.MsoAcetate
	{mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-link:"Balloon Text Char";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:9.0pt;
	font-family:"Segoe UI","sans-serif";
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;}
span.BalloonTextChar
	{mso-style-name:"Balloon Text Char";
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"Balloon Text";
	mso-ansi-font-size:9.0pt;
	mso-bidi-font-size:9.0pt;
	font-family:"Segoe UI","sans-serif";
	mso-ascii-font-family:"Segoe UI";
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:"Segoe UI";
	mso-bidi-font-family:"Segoe UI";}
span.z-TopofFormChar
	{mso-style-name:"z-Top of Form Char";
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"z-Top of Form";
	mso-ansi-font-size:8.0pt;
	mso-bidi-font-size:8.0pt;
	font-family:"Arial","sans-serif";
	mso-ascii-font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Arial;
	mso-bidi-font-family:Arial;
	display:none;
	mso-hide:all;}
span.z-BottomofFormChar
	{mso-style-name:"z-Bottom of Form Char";
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-style-locked:yes;
	mso-style-link:"z-Bottom of Form";
	mso-ansi-font-size:8.0pt;
	mso-bidi-font-size:8.0pt;
	font-family:"Arial","sans-serif";
	mso-ascii-font-family:Arial;
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Arial;
	mso-bidi-font-family:Arial;
	display:none;
	mso-hide:all;}
span.SpellE
	{mso-style-name:"";
	mso-spl-e:yes;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-size:10.0pt;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:.5in .7in .5in 1.0in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
</head>

<body lang=EN-US style='tab-interval:.5in'>

<%
	'Setting the content type to word document.
	Response.ContentType = "application/vnd.ms-doc"
	Response.AddHeader "Content-Disposition", "attachment;filename=Labels.doc" 
%>

<%
	'Calling a function from the .inc file that sets the sql. Then executing the sql into a 
	'record set(rs).
	setSQL()
	Dim SQL
	SQL = Application("SQL")
	Dim rs
	Set rs = connection.execute(SQL)
%>

<div class=WordSection1>

<form>

<table class=MsoNormalTable border=0 cellpadding=0 style='mso-cellspacing:1.5pt;
 background:white;mso-yfti-tbllook:1184;mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <%While Not rs.EOF %>
 <% 
 	'This is checking if the zip is not filled in and setting a variable pertaining to the zip. This also
	'puts some of the record set into asp variables
 	If Not rs.EOF Then
		If rs("Zip")=null Or rs("Zip")="" Or rs("Zip")=0 Then
			zipTemp = ""
		else
			zipTemp = rs("Zip")
		End If
		names = rs("Parfirst") & " " & rs("Parlast")
		city = rs("City") & ", " & rs("State") & " " & zipTemp
 %>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:.95in'>
  <td style='padding:.75pt .75pt .75pt .75pt;height:.95in'>
  <p class=MsoNormal><span style='font-size:11.0pt;mso-fareast-font-family:
  "Times New Roman"'><%=names%><br><%=rs("Street")%><br><%=city%><o:p></o:p></span></p>
  </td>
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
 <td style='padding:.75pt .75pt .75pt .75pt;height:.95in'>
  <p class=MsoNormal style='margin-top:0in;margin-right:.6in;margin-bottom:
  0in;margin-left:.7in;margin-bottom:.0001pt'><span style='font-size:11.0pt;
  mso-fareast-font-family:"Times New Roman"'><%=names%><br><%=rs("Street")%><br><%=city%><o:p></o:p></span></p>
  </td>
  <%  End If%>
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
<td style='padding:.75pt .75pt .75pt .75pt;height:.95in'>
  <p class=MsoNormal style='margin-left:.3in'><span style='font-size:11.0pt;
  mso-fareast-font-family:"Times New Roman"'><%=names%><br><%=rs("Street")%><br><%=city%><o:p></o:p></span></p>
  </td>
  <%  
	End If
	If Not rs.EOF Then
		rs.movenext
	End If
  %>
 </tr>
<%wend%>
</table>

</form>

<%
	rs.close
	Set rs = Nothing 
%>
</div>

</body>

</html>
