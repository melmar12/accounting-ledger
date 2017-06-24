<%

m1 = Request.form("m1")
m2 = Request.form("m2")
m3 = Request.form("m3")
m4 = Request.form("m4")
desc = Request.form("desc")


set rs = Server.CreateObject("ADODB.Recordset")
SQLString = "SELECT * FROM glmaster WHERE "
fsw = 0
if len(m1) > 0 then
   SQLString = SQLString + " major LIKE '" + cstr(m1) + "%'"
   fsw = 1
end if

if len(m2) > 0 and fsw = 0 then
   SQLString = SQLString + " minor LIKE '" + cstr(m2) + "%'"
   fsw = 1
else
   if len(m2) > 0 then
      SQLString = SQLString + " AND minor LIKE '" + cstr(m2) + "%'"
   end if
end if

if len(m3) > 0 and fsw = 0 then
   SQLString = SQLString + " sub1 LIKE '" + cstr(m3) + "%'"
   fsw = 1
else
   if len(m3) > 0 then
      SQLString = SQLString + " AND sub1 LIKE '" + cstr(m3) + "%'"
   end if
end if

if len(m4) > 0 and fsw = 0 then
   SQLString = SQLString + " sub2 LIKE '" + cstr(m4) + "%'"
   fsw = 1
else
   if len(m4) > 0 then
      SQLString = SQLString + " AND sub2 LIKE '" + cstr(m4) + "%'"
   end if
end if


if len(desc) > 0 and fsw = 0 then
   SQLString = SQLString + " acctdesc LIKE '" + cstr(desc) + "%'"
   fsw = 1
else
   if len(desc) > 0 then
      SQLString = SQLString + " AND acctdesc LIKE '" + cstr(desc) + "%'"
   end if
end if

if len(m1) = 0 and len(m2) = 0 and len(m3 ) = 0 and len(m4) = 0 and len(desc) = 0 then
   SQLString = "Select * from glmaster"
   fsw = 1
end if

response.write "<p><table><tr><td colspan='6' bgcolor='#aaaaaa' align='center'><b>Results</td></tr><tr>"
response.write "<td align='center'>Major</td>"
response.write "<td align='center'>Minor</td>"
response.write "<td align='center'>Sub1</td>"
response.write "<td align='center'>Sub2</td>"
response.write "<td align='center'>Account Description</td>"
response.write "<td align='center'>Account<br>Balance</td></tr>"

'response.write "<p>SQL STRING="+cstr(SQLString)

if fsw = 1 then

c = 0
rs.open SQLString,"DSN=gl1425;UID=gl1425;PWD=UGS42ahvG;"
while not rs.eof
    response.write "<tr><td align='center'>"
    response.write cstr(rs("major")) + "</td><td align='right'>"
    response.write cstr(rs("minor")) + "</td><td align='right'>"
    response.write cstr(rs("sub1")) + "</td><td align='right'>"
    response.write cstr(rs("sub2")) + "</td><td align='right'>"
    response.write cstr(rs("acctdesc")) + "</td><td align='right'>"
    response.write formatnumber(rs("balance"),2) + "</td></tr>"
    c = c + 1
rs.movenext
wend
rs.close
set rs = nothing

end if
response.write "</table><p><b>" + cstr(c) + " matching records found"
response.write "<p>SQL STRING=" + cstr(SQLString)
%>




