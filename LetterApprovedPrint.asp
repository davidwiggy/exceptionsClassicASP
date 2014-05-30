<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LetterApprovedPrint" %>
<!--
Developer: David Wiggins
Date:      March 2014
Purpose:   To build a printable word document. The user calls this page to print all the approved
           exceptions for the selected date. This page then formats the layout for a word document.
-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<head>
<meta http-equiv=Content-Type content="text/html; charset=us-ascii">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">

<!--This is declaring the content type as a word document.-->
<%
	Response.ContentType = "application/vnd.ms-doc"
	Response.AddHeader "Content-Disposition", "attachment;filename=Approved.doc" 
%>

<!--Setting the sequel and then executing the sql into a record set-->
<%
	setSQL()
	Dim SQL
	SQL = Application("SQL")
	Dim rs
	Set rs = connection.execute(SQL)
%>

<!--CSS for only this page-->

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
p
	{mso-style-noshow:yes;
	mso-style-priority:99;
	mso-margin-top-alt:auto;
	margin-right:0in;
	mso-margin-bottom-alt:auto;
	margin-left:0in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman","serif";
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;}
span.SpellE
	{mso-style-name:"";
	mso-spl-e:yes;}
span.GramE
	{mso-style-name:"";
	mso-gram-e:yes;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-size:10.0pt;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.3in 1.0in 1.0in 1.0in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
</head>
<!--Beginning to build the layout for word in a While not EOF loop, so that the program will
	run until all the approved exceptions for the selected date are added.-->
<body lang=EN-US style='tab-interval:.5in'>

<div class=WordSection1>

<table class=MsoNormalTable border=0 cellpadding=0 style='mso-cellspacing:1.5pt;
 mso-yfti-tbllook:1184'>
 <% While Not rs.EOF %>
 
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><%=rs("Info data entry date")%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><br>
  <%=rs("Parfirst")%><%=(" ")%><%=rs("Parlast")%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><%=rs("Street")%>
  <span class=SpellE>Trescott</span> Ct.<o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><%=rs("City")%><%If rs("City") <> "" Then Response.Write(", ") End If %><%=rs("State")%><%=(" ")%><%if rs("Zip") <> 0 Then 
																			                                           Response.Write(rs("Zip")) 
																	                                                End If%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><br>
  <%=rs("Stufirst")%><%=(" ")%><%=rs("Stulast")%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span class=SpellE><span style='mso-fareast-font-family:
  "Times New Roman"'><%=rs("SchoolName")%></span></span>
  </td>
 </tr>
 <tr style='mso-yfti-irow:6'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal style='margin-bottom:12.0pt'><span style='mso-fareast-font-family:
  "Times New Roman"'>Approval for Exceptions: <%=rs("Type of exception")%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:7'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p>Dear Sir/Madame<span class=GramE>:</span><br>
  <br>
  Your request for an exception to the above mentioned
  child's/children's regular geographic school assignment has been
  given careful administrative review.<br>
  <br>
  After thorough consideration of all evidence and documentation presented in
  support of this petition, your request has been APPROVED as being in
  compliance with the recognized categories of exceptions authorized by
  Berkeley County School Board policies.<br>
  <br>
  This approval is only for the <%=rs("Schoolyear")%> school year. Request for exceptions
  must be made on an annual basis through my office.<br>
  <br>
  If we can answer any further questions or assist you in any way, please do
  not hesitate to call.<br>
  <br>
  Sincerely, <br>
  <br>
  <br>
  <br>
  Charlie Davis, Administrative Assistant for Superintendent <br>
  Division of Administration and Pupil Services <br>
  <br>
  BD/<span class=SpellE>ps</span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><br>
  C: <%=rs("SchoolPrincipal")%><o:p></o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:9;mso-yfti-lastrow:yes'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'>File<o:p></o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  <p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <%  If Not rs.EOF Then 
		rs.movenext 
	  End If
  %>
 </tr>
 <% wend %>
</table>

<p class=MsoNormal><span style='mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>

</div>

</body>
</HTML>