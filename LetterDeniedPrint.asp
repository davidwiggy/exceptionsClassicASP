<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% DBLink = "ExceptionsIndex" %>
<!--#include virtual ="connections.asp"-->
<!--#include file="sql.inc"-->
<!--Setting the global varible for the current page-->
<% Application("page")="LetterDeniedPrint" %>

<!--
	Developer: David Wiggins
    Date: March 2014
    Purpose: The purpose of this file is to allow the user to print the denied letters.
             
-->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>
<head>
 
<%
	'This is declaring the content type of word document.
	Response.ContentType = "application/vnd.ms-doc"
	Response.AddHeader "Content-Disposition", "attachment;filename=Denied.doc" 
%>

<%
	'This calling a function from the .inc file to set the SQL. Then executes the sql and puts the result into 
	'a record set(rs).
	setSQL()
	Dim SQL
	SQL = Application("SQL")
	Dim rs
	Set rs = connection.execute(SQL)
%>
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
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-536870145 1073786111 1 0 415 0;}
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
	font-family:"Calibri","sans-serif";
	mso-ascii-font-family:Calibri;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:"Times New Roman";
	mso-fareast-theme-font:minor-fareast;
	mso-hansi-font-family:Calibri;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;}
.MsoPapDefault
	{mso-style-type:export-only;
	margin-bottom:8.0pt;
	line-height:107%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in .5in 1.0in 1.0in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
</head>
<!-- Building the Main Menu*************************************************************-->
<body lang=EN-US style='tab-interval:.5in'>

<div class=WordSection1>

<table class=MsoNormalTable border=0 cellpadding=0 style='mso-cellspacing:1.5pt;
 mso-yfti-tbllook:1184'>
 <% While Not rs.EOF %>
 
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpFirst><%=rs("Info data entry date")%></p><br />
  </td>
 </tr>
 <tr style='mso-yfti-irow:1'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><%=rs("Parfirst")%><%=(" ")%><%=rs("Parlast")%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:2'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><%=rs("Street")%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:3'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><%=rs("City")%><%If rs("City") <> "" Then Response.Write(", ") End If %><%=rs("State")%><%=(" ")%><%if rs("Zip") <> 0 Then 
																			                                           Response.Write(rs("Zip")) 
																	                                                End If%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:4'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><br>
  <%=rs("Stufirst")%><%=(" ")%><%=rs("Stulast")%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><span class=SpellE><%=rs("SchoolName")%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:6'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle>Approval for Exceptions: <%=rs("Type of exception")%> </p><br />
  </td>
 </tr>
 <tr style='mso-yfti-irow:7'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p style='margin-bottom:0in;margin-bottom:.0001pt;mso-add-space:auto'>Dear
  Sir/Madame<span class=GramE>:</span><br>
  <br>
  Your request for an exception to your child’s/children’s regular geographic
  school assignment has been given careful administrative review.<br>
  <br>
  After thorough consideration of all evidence and documentation presented in
  support of this petition, I must regretfully deny your request since it does
  not comply with any of the recognized categories of exceptions approved by
  Berkeley County School Board policies.<br>
  <br>
  Should you wish to appeal this decision, you may make a written request to be
  heard by the Berkeley County Board of Education at a regular session. Your
  request should be made within ten (10) days from the date of this letter and
  should be addressed to me or to the Board Chairperson. Such petitions may be
  heard at the discretion of the Board. If the Board agrees to consider your
  petition, you will be notified of the date and time of the hearing.<br>
  <br>
  If we can answer any further questions or assist you in any way, please do
  not hesitate to call.<br>
  <br>
  Sincerely,</p>
  <p style='margin-bottom:0in;margin-bottom:.0001pt;mso-add-space:auto'><br
  style='mso-special-character:line-break'>
  <![if !supportLineBreakNewLine]><br style='mso-special-character:line-break'>
  <![endif]></p>
  <p style='margin-bottom:0in;margin-bottom:.0001pt;mso-add-space:auto'><br>
  Charlie Davis, Administrative Assistant for Superintendent </p>
  <p style='margin-bottom:0in;margin-bottom:.0001pt;mso-add-space:auto'>Division
  of Administration and Pupil Services<br>
  <br>
  CD/<span class=SpellE>ps</span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><br>
  C: <%=rs("SchoolPrincipal")%></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:9;mso-yfti-lastrow:yes'>
  <td style='padding:.75pt .75pt .75pt .75pt'>
  <p class=MsoNormalCxSpMiddle><span style='mso-spacerun:yes'>     </span>File</p><br /><br /><br /><br /><br /><br />
  </td>
 <%  If Not rs.EOF Then 
		rs.movenext 
	  End If
  %>
 </tr>
 <% wend %>
</table>

<p class=MsoNormal style='margin-bottom:8.0pt;line-height:107%'><o:p>&nbsp;</o:p></p>

</div>

</body>

</html>

