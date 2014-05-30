<% 'Setup Root Directories
	Dim root
	root = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\inetpub\wwwroot\Exceptionsdb.accdb"
	'Create Connections
	Dim connection
	Set connection = Server.CreateObject("ADODB.Connection")
	connection.Open root
%>