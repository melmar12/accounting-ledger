<html>
    <head>
    <!-- meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">
    <!-- custom css -->
    <link rel="stylesheet" href="styles.css">
<body>
  <div class="container-fluid">

	      <div class="row justify-content-md-center navi">
	        <div class="col-11">
	          <a href="./home.html">home</a> |
            <a href="./"> index</a>
	        </div>
	      </div>

	    <div class="row justify-content-md-center">
	        <div class="col-8 col-md-auto">

				<h2>General Ledger Report</h2>
				<%
				Set rs = Server.CreateObject("ADODB.Recordset")
				sql_string="Select * from je ORDER BY sourceref ASC, srseq ASC"

				'response.write "<p>SQL--->"+sql_string+"<---<p>"
				rs.open sql_string, "DSN=gl1425;UID=gl1425;PWD=UGS42ahvG;"

				response.write "<p><table><tr>"

				for i = 0 to rs.fields.count - 1
				  response.write "<td align='center'>"+cstr(rs(i).Name)+"</td>"
				next
				response.write "</tr>"
				row_count=0
				while not rs.EOF
				  row_count=row_count+1
				  response.write "<tr>"
				  for i = 0 to rs.fields.count - 1
				     response.write "<td align='right'>"+cstr(rs(i))+"</td>"
				  next
				  response.write "</tr>"
				  rs.MoveNext
				wend

				response.write "</table><p>Row Count="+cstr(row_count)
				response.write "<p>Normal Termination "+cstr(now)
				rs.Close
				set rs=nothing
				%>
    </div>
   </div>
  </div>

    <!-- bootstrap v4 alpha JS -->
  <script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/tether/1.4.0/js/tether.min.js" integrity="sha384-DztdAPBWPRXSA/3eYEEUWrWCy7G5KFbe8fFjk5JAIxUYHKkDx6Qin1DkWx51bBrb" crossorigin="anonymous"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js" integrity="sha384-vBWWzlZJ8ea9aCX4pEW3rVHjgjt7zpkNpZk+02D9phzyeVkE+jo0ieGizqPLForn" crossorigin="anonymous"></script>
</body>
</HTML>
