<!DOCTYPE html>
<html lang="en">
  <head>
    <title>Delete an Account</title>
    <!-- meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">
    <!-- custom css -->
    <link rel="stylesheet" href="styles.css"/>
    <style media="screen">
      .navi {
        border-bottom: 1px solid #eee;
      }

      h1, h2 {
          text-align: center;
          padding: 50px 0px;
      }

      td {
          padding: 5px;
          border: 1px solid grey;
      }
      .btn {
        margin-top: 20px;
        float: right;
      }
    </style>
  </head>
<body>
    <div class="container-fluid">

      <div class="row justify-content-md-center navi">
        <div class="col-11">
          <a href="./">index</a>
          <a href="./home.html">home</a>
        </div>
      </div>

      <div class="row justify-content-md-center">
          <div class="col-3">

        		<%
        		sub pass1
        		%>
            <h1>Delete Account</h1>
        		<form action="a43.asp" method="POST" name="f1">
              <table>
                  <tr>
                      <td>Major</td>
                      <td>Minor</td>
                      <td>Sub-1</td>
                      <td>Sub-2</td>
                  </tr>
                  <tr>
                      <td><input class="form-control" type="text" size="8" name="major"></td>
                      <td><input class="form-control" type="text" size="8" name="minor"></td>
                      <td><input class="form-control" type="text" size="8" name="sub1" value="0"></td>
                      <td><input class="form-control" type="text" size="8" name="sub2" value="0"></td>
                  </tr>
              </table>
              <input type="hidden" name="token" value="2">
              <div class="form-group row">
                <div class="col-12">
                  <button type="submit" class="btn btn-danger">Delete Account</button>
                </div>
              </div>
        		</form>
          	<%
          	end sub

          	sub pass2

          	set cn = Server.CreateObject("ADODB.Connection")
          	cn.open "gl1425","gl1425","UGS42ahvG"
          	response.write "<P>Connection created OK"

          	  set rs = Server.CreateObject("ADODB.Recordset")
          	  major = request.form("major")
          	  minor = request.form("minor")
          	  sub1 = request.form("sub1")
          	  sub2 = request.form("sub2")

          	  SQLString = "SELECT * FROM glmaster WHERE major="+cstr(major)+" AND minor="+cstr(minor)+" AND sub1="+cstr(sub1)+" AND sub2="+cstr(sub2)

          	  response.write "<P>SQL</br>"+cstr(SQLString)

          	  rs.open SQLString,"DSN=gl1425;UID=gl1425;PWD=UGS42ahvG;"
          	  response.write "<P>Recordset opened OK"

          	   c=0
          	   while NOT rs.EOF
          	      c=c+1
          	      bal = rs("balance")
          	    if not (cdbl(bal) = 0) then
          	        c=99
          	    end if
          	    rs.movenext
          	   wend

          	   response.write "<P>Looking for for existing account: "
          	        if c=1  then
          	           response.write ": Found "+cstr(c)+" records. </br> Current balance is: " + cstr(bal)+". Proceeding with delete"
          	            SQLString = "DELETE FROM glmaster WHERE major="+cstr(major)+" AND minor="+cstr(minor)+" AND sub1="+cstr(sub1)+" AND sub2="+cstr(sub2)

          	           response.write "<p>ready to delete SQL: " + SQLString

          	           cn.execute SQLString,numa
          	           if numa=1 then
          	                response.write "<P>Deleted "+cstr(numa) + " row"
          	           else
          	                response.write "<P>delete Failed. Number of records deleted="+cstr(numa)
          	           end if
          	           cn.close
          	           set cn=nothing

          	        else
          	           if c=0 then
          	                response.write ": Found "+cstr(c)+" records. Delete will be not performed. "
          	           else if c=99 then
          	                response.write "Found account.</br> Current balance is: " + cstr(bal)+", delete failed"
          	           end if
          	           end if
          	        end if

          	    rs.close
          	    set rs=nothing
          	end sub

          	sub passerror
          	     response.write "<p>INVALID TOKEN VALUE. token="+cstr(tokernvalue)
          	end sub


          	tokenvalue=request.form("token")
          	select case tokenvalue
          	case ""
          	   call pass1
          	case "2"
          	  call pass2
          	case else
          	   call passerror
          	end select
          	%>



    </div>
  </div>
</div>
  <!-- bootstrap v4 alpha JS -->
  <script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/tether/1.4.0/js/tether.min.js" integrity="sha384-DztdAPBWPRXSA/3eYEEUWrWCy7G5KFbe8fFjk5JAIxUYHKkDx6Qin1DkWx51bBrb" crossorigin="anonymous"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js" integrity="sha384-vBWWzlZJ8ea9aCX4pEW3rVHjgjt7zpkNpZk+02D9phzyeVkE+jo0ieGizqPLForn" crossorigin="anonymous"></script>
</body>
</html>
