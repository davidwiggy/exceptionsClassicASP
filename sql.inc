<%
 	sub setSQL()
		Select Case Application("page")
		
			Case "ExceptionsIndex"
				Application("SQL") = "SELECT SchoolYear FROM dbo_tblYearDates Order By SchoolYear DESC"
				
			Case "ExceptionsList"
				Application("SQL") = "SELECT schoolyear, race, [Info data entry date], Approved, SASIId, Id, Stulast," & _
									 "StuFirst, [Sending Sch], ReceivingSch,[Type of Exception] FROM dbo_AllExceptions " & _
									 "WHERE schoolyear='" & year & "' ORDER BY schoolyear DESC, stulast, stufirst"
									 
			Case "SdntDetail"
				id = application("studentId")
				Application("SQL") = "SELECT * FROM dbo_AllExceptions WHERE Id =" & id & ""
									 
			Case "SdntDetail2"
				Application("SQL") = "SELECT DISTINCT [SchoolInit] FROM dbo_tblSchools"
				
			Case "SdntDetail3"
				Application("SQL") = "SELECT DISTINCT [Type] FROM dbo_tblExceptionTypes ORDER BY [type]"
			
			Case "SdntDetail4"
				Application("SQL") = "SELECT SchoolYear FROM dbo_tblYearDates Order By SchoolYear DESC"
				
			Case "SdntDetailUpdate"
				Dim currentDate
				currentDate = Date()
				Application("SQL") = "UPDATE dbo_AllExceptions "  &_
					"SET SchoolYear		 	='" & Request.Form("year")         & "', " &_
					"ReceivingSch  		 	='" & Request.Form("receivingSch") & "', " &_
					"StuFirst      		 	='" & Request.Form("stuFirst")     & "', " &_
					"Stulast       		 	='" & Request.Form("stuLast")      & "', " &_
					"Race          		 	='" & Request.Form("race")         & "', " &_
					"Grade         		 	='" & Request.Form("grade")        & "', " &_
					"ParFirst      		 	='" & Request.Form("parFirst")     & "', " &_
					"Parlast       		 	='" & Request.Form("parLast")      & "', " &_
					"Street        		 	='" & Request.Form("address")      & "', " &_
					"City          		 	='" & Request.Form("city")         & "', " &_
					"State         		 	='" & Request.Form("state")        & "', " &_
					"Zip           		 	="  & Request.Form("zip")          & ",  " &_
					"[Sending Sch]       	='" & Request.Form("sendingSch")   & "', " &_
					"Approved               ="  & Request.Form("approved")     & ",  " &_
					"[Type of Exception]    ='" & Request.Form("exception")    & "', " &_
					"[Info data entry date] ='" & Request.Form("entryDate")    & "', " &_
					"SASIId                 ='" & Request.Form("sas")          & "'  " &_
					"WHERE Id=" & Application("studentId")
				
			Case "AddSdntExceptions"
				Application("SQL") = "INSERT INTO dbo_AllExceptions " &_
					"(SASIId, Stulast, [Sending Sch], StuFirst, ReceivingSch, Parlast, Grade, Parfirst, Race, Street, Zip, City, Approved, State, [Type of Exception], SchoolYear, [Info data entry date])" &_
					"VALUES " &_
					"('" & Request.Form("sas")          & "'," &_
					"'"  & Request.Form("stuLast")      & "'," &_
					"'"  & Request.Form("sendingScho")  & "'," &_
					"'"  & Request.Form("stuFirst")     & "'," &_
					"'"  & Request.Form("receivingSch") & "'," &_
					"'"  & Request.Form("parLast")      & "'," &_
					"'"  & Request.Form("grade")        & "'," &_
					"'"  & Request.Form("parFirst")     & "'," &_
					"'"  & race                         & "'," &_
					"'"  & Request.Form("address")      & "'," &_
					" "  & Request.Form("zip")          & ", " &_
					"'"  & Request.Form("city")         & "'," &_
					" "  & Request.Form("approved")     & ", " &_
					"'"  & Request.Form("state")        & "'," &_
					"'"  & Request.Form("exception")    & "'," &_
					"'"  & Request.Form("year")         & "'," &_
					"'"  & Request.Form("entryDate")    & "');"
					
			Case "SdntInserted"
				Application("SQL") = "SELECT Id FROM dbo_AllExceptions " &_
									 "WHERE  Stulast='" & Request.Form("stuLast")  & "'" &_
									 "AND   StuFirst='" & Request.Form("stuFirst") & "'" &_
									 "AND    ParLast='" & Request.Form("parLast")  & "'" &_
									 "AND   ParFirst='" & Request.Form("parFirst") & "'" &_
									 "AND     Street='" & Request.Form("address")  & "'" &_
									 "AND        Zip="  & Request.Form("zip")      & " " &_
									 "AND       Race='" & race                     & "'" &_
									 "AND SchoolYear='" & Request.Form("year")     & "'"
									 
			Case "LettersAndLabels"
				Application("SQL") = "SELECT DISTINCT [Info data entry date] FROM dbo_AllExceptions ORDER BY [Info data entry date]"
										 
			Case "LabelsExceptions"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
			Case "LetterApproved"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.Approved <> 0 AND dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
			Case "LetterDenied"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.Approved = 0 AND dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
			Case "ExceptionTypes"
				Application("SQL") = "SELECT * FROM dbo_tblExceptionTypes"
				
			Case "ReceivingSchool"
				Application("SQL") = "SELECT DISTINCT ReceivingSch FROM dbo_AllExceptions ORDER BY ReceivingSch"
				
			Case "ReceivingSchool2"
				Application("SQL") = "SELECT TOP 100 PERCENT * FROM dbo_AllExceptions " &_
									 "WHERE ReceivingSch='" & school & "'" & " AND SchoolYear='" & schYear & "' ORDER BY Stulast"
									 
			Case "LabelsPrint"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
			Case "LetterApprovedPrint"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.Approved <> 0 AND dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
			Case "LetterDeniedPrint"
				Application("SQL") = "SELECT dbo_AllExceptions.*, dbo_tblSchools.SchoolName, dbo_tblSchools.SchoolPrincipal " &_
									 "FROM dbo_AllExceptions LEFT JOIN dbo_tblSchools ON dbo_AllExceptions.ReceivingSch = dbo_tblSchools.SchoolInit " &_
									 "WHERE dbo_AllExceptions.Approved = 0 AND dbo_AllExceptions.[Info data entry date]=#" & Application("labelDate") & "#"
									 
		End Select
	
	
	End sub
		
%>
