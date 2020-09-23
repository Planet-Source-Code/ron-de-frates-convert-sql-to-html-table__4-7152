<div align="center">

## Convert SQL To HTML Table


</div>

### Description

Returns the results of a SQL query as an HTML table. This small function is not only very fast it also simplifies the programming logic for displaying data.
 
### More Info
 
SQL query

Requires ADO 2.0 or higher. I benchmarked this code against a traditional "rs.MoveNext" method and got a 100% speed improvement for 200 records.

A string containing the results as a HTML table


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ron de Frates](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ron-de-frates.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ron-de-frates-convert-sql-to-html-table__4-7152/archive/master.zip)





### Source Code

```
<%
' ----------------------------------------------------------------------------
' NAME:    cnvrtRSToHTML
' AUTHOR:   Ron de Frates
' DATE:    01/21/2002
' PURPOSE:   Converts recordset results to HTML table
' LANGUAGE:  ASP, HTML
' OBJECTS:   ADODB.Connection
' DATABASES:  DSN_Admin
' ----------------------------------------------------------------------------
Function cnvrtRSToHTML(sSQL)
	Dim dbConn, rsSrc, sHTMLTbl
	Set dbConn = Server.CreateObject("ADODB.Connection")
	dbConn.Open
	Set rsSrc = dbConn.Execute(sSQL)
	sHTMLTbl = rsSrc.GetString(,,"</td><td>","</td></tr><tr><td>","&nbsp;")
	sHTMLTbl = Mid(sHTMLTbl, 1, Len(sHTMLTbl) - 8)
	sHTMLTbl = "<table border=1><tr><td>" & sHTMLTbl & "</table>"
	cnvrtRSToHTML = sHTMLTbl
End Function
' -----------------------------------------------------------------------------
' ***** HOW TO CALL THIS FUNCTION *****
Dim sSQL
sSQL = "SELECT * FROM myTable "
Response.Write(cnvrtRSToHTML(sSQL))
%>
```

