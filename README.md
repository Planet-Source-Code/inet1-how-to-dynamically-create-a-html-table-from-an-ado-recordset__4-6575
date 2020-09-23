<div align="center">

## How to dynamically create a HTML table from an ADO recordset \!\.


</div>

### Description

Shows how to dynamically create a HTML table from a recordset. All you need is the connection string and the table name.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[iNet1](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/inet1.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/inet1-how-to-dynamically-create-a-html-table-from-an-ado-recordset__4-6575/archive/master.zip)





### Source Code

```
<XMP>
**********************************************************************
	How to dynamically create a html table from an ADO recordset !.
	All you need is your connection string and Table name, and
	off you go !.
**********************************************************************
<%
	Dim conn
	Dim cmd
	Dim rs
	Set conn = Server.CreateObject("ADODB.Connection")
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = "Provider = MSDASQL;Data Source=Northwind;Database=Northwind;User Id=;Password=;"
	cmd.CommandText = "SELECT * FROM Customers"
	Set rs = cmd.Execute
	Response.Write "<table width=100% border=1>"
	Response.Write "<tr>"
	for i = 1 to rs.Fields.Count - 1
		Response.Write "<td><strong>" & rs.Fields(i).Name & "<strong></td>"
	next
		Response.Write "</tr>"
	Do While Not rs.EOF
			Response.Write "<tr>"
		for i = 1 to rs.Fields.Count - 1
				Response.Write "<td>" & rs.Fields(i) & "</td>"
		next
		rs.MoveNext
			Response.Write "</tr>"
	Loop
			Response.Write "</table>"
%>
</XMP>
```

