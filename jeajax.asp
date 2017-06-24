<%

'
' jeajax.asp
'
'
'   receives form submission from AjaxRequest.submit
'   (program: a46.htm)
'


m1=Request.form("m1")


set rs=Server.CreateObject("ADODB.Recordset")
SQLString="SELECT * FROM je WHERE "
fsw=0
if len(m1) > 0 then
   SQLString=SQLString+" sourceref LIKE '"+cstr(m1) +"%'"
   fsw=1
end if

if len(m1) =0 and len(m2)=0 and len(m3)=0 and len(m4)=0 and len(desc)=0 then
   SQLString="Select * from je"
   fsw=1
end if

response.write "<p><table border='1'><tr><td colspan='8' bgcolor='#aaaaaa' align='center'><b>Results</td></tr><tr>"
response.write "<td align='center'>sourceref</td>"
response.write "<td align='center'>srseq</td>"
response.write "<td align='center'>jemajor</td>"
response.write "<td align='center'>jeminor</td>"
response.write "<td align='center'>jesub1</td>"
response.write "<td align='center'>jesub2</td>"
response.write "<td align='center'>jedesc</td>"
response.write "<td align='center'>jeamount</td></tr>"

if fsw=1 then

c=0
rs.open SQLString,"DSN=gl1425;UID=gl1425;PWD=UGS42ahvG;"
while not rs.eof
    
    response.write "<tr><td align='center'>"
    response.write cstr(rs("sourceref"))+"</td><td align='right'>"
    response.write cstr(rs("srseq"))+"</td><td align='right'>"
    response.write cstr(rs("jemajor"))+"</td><td align='right'>"
    response.write cstr(rs("jeminor"))+"</td><td align='right'>"
    response.write cstr(rs("jesub1"))+"</td><td align='right'>"
    response.write cstr(rs("jesub2"))+"</td><td align='right'>"
    response.write rs("jedesc")+"</td><td align='right'>"
    response.write cstr(rs("jeamount"))+"</td></tr>"
    c=c+1
rs.movenext
wend
rs.close
set rs=nothing

end if
response.write "</table><p><b>"+cstr(c)+" matching records found"
'response.write "<p>SQL STRING="+cstr(SQLString)


%>




