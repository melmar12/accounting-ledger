<html>
  <head>
      <title>Add Account</title>
    <!-- meta tags -->
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap CSS -->
    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/css/bootstrap.min.css" integrity="sha384-rwoIResjU2yc3z8GV/NPeZWAv56rSmLldC3R/AZzGRnGxQQKnKkoFVhFQhNUwEyJ" crossorigin="anonymous">
    <!-- custom css -->
    <link rel="stylesheet" type="text/css" href="styles.css">
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
            <div class="col-6">

                    <%
                    sub pass1
                    %>

                    <h1>Add Account to General Ledger</h1>
                    <form name="xform" method="POST" action="http://auckland.bauer.uh.edu/students/gl1425/a41.asp">
                        <table>
                            <tr>
                                <td>Major</td>
                                <td>Minor</td>
                                <td>Sub-1</td>
                                <td>Sub-2</td>
                                <td>Description</td>
                            </tr>
                            <tr>
                                <td><input class="form-control" type="text" size="8" name="major"></td>
                                <td><input class="form-control" type="text" size="8" name="minor"></td>
                                <td><input class="form-control" type="text" size="8" name="sub1"></td>
                                <td><input class="form-control" type="text" size="8" name="sub2"></td>
                                <td><input class="form-control" type="text" size="40" name="acctdesc"></td>
                            </tr>
                        </table>

                        <input type="hidden" name="token" value="2">
                        <div class="form-group row">
                          <div class="col-12">
                            <button type="submit" class="btn btn-primary">Add Account</button>
                          </div>
                        </div>

                    <%
                    end sub
                    sub pass2

                    set cn = Server.CreateObject("ADODB.Connection")
                    fdsn="gl1425"
                    fuid="gl1425"
                    fpwd="UGS42ahvG"
                    cn.open fdsn,fuid,fpwd

                    response.write "open ok..."

                    major = Request.form("major")
                    minor = Request.form("minor")
                    sub1 = Request.form("sub1")
                    sub2 = Request.form("sub2")
                    acctdesc = Request.form("acctdesc")
                    balance = 0.0

                    Insert_string="INSERT INTO glmaster (major,minor,sub1,sub2,acctdesc,balance) VALUES ("+cstr(major)+","+cstr(minor)+","+cstr(sub1)+","+cstr(sub2)+","+chr(39)+cstr(acctdesc)+chr(39)+", "+cstr(balance)+")"

                    response.write "<p>insert_string="+cstr(Insert_string)

                    cn.execute Insert_string,numa

                    if numa=1 then
                         response.write "<p>Success!"
                    else
                      response.write "Error :("
                    end if
                    end sub


                    sub errorpass
                      response.write "errorpass"
                    end sub

                    token=request.form("token")
                    select case token
                    case ""
                       call pass1
                    case "2"
                      call pass2
                    case else
                      call errorpass
                    end select
                    %>
                    </form>

                </div>
            </div>
        </div>

            <!-- bootstrap v4 alpha JS -->
  <script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/tether/1.4.0/js/tether.min.js" integrity="sha384-DztdAPBWPRXSA/3eYEEUWrWCy7G5KFbe8fFjk5JAIxUYHKkDx6Qin1DkWx51bBrb" crossorigin="anonymous"></script>
  <script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js" integrity="sha384-vBWWzlZJ8ea9aCX4pEW3rVHjgjt7zpkNpZk+02D9phzyeVkE+jo0ieGizqPLForn" crossorigin="anonymous"></script>
  </body>
</html>
