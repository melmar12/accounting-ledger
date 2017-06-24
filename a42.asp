<html>
    <title>Update an Account</title>
  <head>
    <!-- meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">
    <!-- custom css -->
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
          <a href="./">index</a> |
          <a href="./home.html">home</a>
        </div>
      </div>

      <div class="row justify-content-md-center">



      <%
      sub pass1
      %>
      <div class="col-3">
      <h1>Find Account</h1>
      <form action="a42.asp" method="POST" name="f1">
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
            <button type="submit" class="btn btn-primary">Find Account</button>
          </div>
        </div>
      </form>
    </div>
      <%
      end sub

      sub pass2
      %>
      <div class="col-6">
      <%

      set rs=server.createobject("ADODB.recordset")

      sql="SELECT * FROM glmaster WHERE major="+request.form("major")
      sql=sql+" AND minor="+request.form("minor")
      sql=sql+" AND sub1="+request.form("sub1")
      sql=sql+" AND sub2="+request.form("sub2")

      'response.write "<p>SQL="+sql+"</p>"

      rs.open sql,"DSN=gl1425;UID=gl1425;PWD=UGS42ahvG;"
      response.write "<p>Recordset opened OK</p></br>"

      c=0
      while NOT rs.eof
          ad=rs("acctdesc")
          c=c+1
          rs.movenext
      wend
      rs.close
      set rs=nothing

      if c=0 then
        response.write "No records found"
      else
        response.write "Record found"
      %>
      <h1>Update Account</h1>
      <form action=a42.asp method="POST" name="f1">

      <input type="hidden" name="major" value="<% =request.form("major") %>">
      <input type="hidden" name="minor" value="<% =request.form("minor") %>">
      <input type="hidden" name="sub1" value="<% =request.form("sub1") %>">
      <input type="hidden" name="sub2" value="<% =request.form("sub2") %>">
      <input type="hidden" name="token" value="3">

      <table>
          <tr>
              <td>Major</td>
              <td>Minor</td>
              <td>Sub-1</td>
              <td>Sub-2</td>
              <td>Description</td>
          </tr>
          <tr>
              <td><% =request.form("major") %></td>
              <td><% =request.form("minor") %></td>
              <td><% =request.form("sub1") %></td>
              <td><% =request.form("sub2") %></td>
              <td><input class="form-control" type="text" size="50" value="<% = ad %>" name="ad"></td>
          </tr>
      </table>

      <div class="form-group row">
        <div class="col-12">
          <button type="submit" class="btn btn-primary">Update Descrption</button>
        </div>
      </div>

      </form>
      </div>

      <%
      end if

      end sub

      sub pass3
      %>
      <div class="col-6">
      <%

      set cn=server.createobject("ADODB.connection")
      cn.open "gl1425","gl1425","UGS42ahvG"
      response.write "</br>Connection opened OK"
      description=Request.form("ad")
      while NOT Instr(cstr(description),"'") = 0
          description=Replace(description,"'"," ")
      wend
      sql ="UPDATE glmaster SET acctdesc="+chr(39)+description+chr(39)
      sql=sql+" WHERE major="+request.form("major") +" AND "
      sql=sql+" minor="+request.form("minor") +" AND "
      sql=sql+" sub1="+request.form("sub1") +" AND "
      sql=sql+" sub2="+request.form("sub2")

      response.write "<p>SQL: "+sql+"</p>"

      cn.execute sql,numa

      if numa=1 then
       response.write "<p>Update successful"
      else
       response.write "<p>UPDATE FAILED"
      end if

      cn.close
      set cn=nothing
      %>
      </div>
      <%
      end sub

      '*****MAIN
      '

      token_value=Request.form("token")
      select case token_value
      case ""
        call pass1
      case "2"
        call pass2
      case "3"
        call pass3
      end select
      %>



  </div>
</div>
  <!-- bootstrap v4 alpha JS -->
  <script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/tether/1.4.0/js/tether.min.js" integrity="sha384-DztdAPBWPRXSA/3eYEEUWrWCy7G5KFbe8fFjk5JAIxUYHKkDx6Qin1DkWx51bBrb" crossorigin="anonymous"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js" integrity="sha384-vBWWzlZJ8ea9aCX4pEW3rVHjgjt7zpkNpZk+02D9phzyeVkE+jo0ieGizqPLForn" crossorigin="anonymous"></script>
</body>
</html>
