<html>
  <head>
  <title>a3</title>
  </head>

  <body>




                <table border="1" valign='middle' bgcolor='#999999'>
                  <tr>
                      <td colspan="3"><center><font color="#ffffff"><br><b>General Ledger</b><br></td>
                  </tr>
                  <tr>
                      <td><center><a href="../index.htm">Index</a></td>
                      <td><a href="a44.asp" target="_blank">Full Trial Balance</a></td>
                      <td><a href="a3.htm">Test Again</a></td>
                  </tr>
              </table></br>Posting journal entry</br></br> 
                  </B>
      <%

      '
      '   (program: a3.htm)
      '

      ' Gather input

      srcnum = Request.Form("sourceref")
      err_count = 0
      what_error = ""
      Dim arrayInput(7,7)
      validsrc = 0
      validacctnum = 0
      balances = Array(0.0,0.0,0.0,0.0,0.0,0.0,0.0)
      compareAccts = Array("","","","","","","")
      for i = 0 to Request.Form("numvalid")
         for j = 0 to 5
              rowHolder = ("c" + cstr(j) + "r" + cstr(i))
              arrayInput(j,i) = Request.Form(rowHolder)
          next
      next
      'for as many entries that are not null, loop and put them into the multidimensional array
      'Existence test for je
      response.write "-----------Checking for valid source refernce number-----------</br>"
      set rs=Server.CreateObject("ADODB.Recordset")
          SQLString="SELECT * FROM je WHERE sourceref =" + srcnum

          rs.open SQLString,"DSN=gl12345;UID=gl12345;PWD=QWERTYYUIOP;"

      if rs.eof then
          response.write "</br>Source Reference Number "+ srcnum +" valid</br>"
          validsrc = 1
      else
          err_count = err_count+1
          what_error = what_error + "</br>Source Reference Number invalid. Transaction cancelled</br>"
      end if
      rs.close
      set rs=nothing




      'Existence test for accounts
      'Col, row

      for k = 0 to Request.Form("numvalid")-1
          set rs=Server.CreateObject("ADODB.Recordset")
          if(IsNumeric(arrayInput(0,k))) then
              SQLString="SELECT * FROM glmaster WHERE major="+arrayInput(0,k)
              SQLString=SQLString+" AND minor="+arrayInput(1,k)
              SQLString=SQLString+" AND sub1="+arrayInput(2,k)
              SQLString=SQLString+" AND sub2="+arrayInput(3,k)
              'response.write "</br>"+ SQLString
              rs.open SQLString,"DSN=gl12345;UID=gl12345;PWD=QWERTYYUIOP;"
              
          else 
              response.write "</br>Finished checking accounts</br>"
              exit for
          end if
          if rs.eof then
              err_count = err_count+1
              what_error =what_error+ "</br>Account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k)+ " not found. </br>"
          end if
          if NOT rs.eof then
              balances(k) = rs("balance")
              response.write "</br>Checking account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k)
              response.write "</br>Account Number valid</br>"
              validacctnum = validacctnum+1
              
          end if

          rs.close
          set rs=nothing
      next

      'Checking for doubles

      for k = 0 to validacctnum
          totalacct = cstr(arrayInput(0,k)) + cstr(arrayInput(1,k)) + cstr(arrayInput(2,k)) + cstr(arrayInput(3,k))
          for j = k+1 to validacctnum
              nextacct = cstr(arrayInput(0,j)) + cstr(arrayInput(1,j)) + cstr(arrayInput(2,j)) + cstr(arrayInput(3,j))
              isDouble = StrComp(totalacct,nextacct)
              if isDouble = 0 then
                  err_count = err_count + 1
                  what_error = what_error + "</br>Accounts can only be listed once.</br>"
                  validacctnum = validacctnum-1
              end if
          next
      next
             
      if validacctnum>1 then  
          response.write "</br>Found " + cstr(validacctnum)+ " valid accounts.</br>"
      else
          response.write "</br>Found " + cstr(validacctnum)+ " valid account.</br>"
      end if
      response.write "Number of errors is " + cstr(err_count) +"</br>"


      'Posting to JE table
      set cn=server.createobject("ADODB.connection")   
      cn.open "gl12345","gl12345","QWERTYYUIOP"
      cn.BeginTrans
      if err_count = 0 then
          response.write "</br>-----------Creating journal entries-----------</br>"
              

          for k = 0 to Request.Form("numvalid")-1

              if NOT arrayInput(5,k) = "" then
                  while NOT Instr(cstr(arrayInput(5,k)),"'") = 0
                      arrayInput(5,k)=Replace(arrayInput(5,k),"'"," ")
                  wend
                  sql ="INSERT INTO je (sourceref, srseq, jemajor, jeminor, jesub1, jesub2, jeamount, jedesc) VALUES ("
                  sql = sql + cstr(srcnum) + ","+cstr(k+1)+ ","+ arrayInput(0,k) + "," + arrayInput(1,k)
                  sql = sql + "," + arrayInput(2,k) + "," + arrayInput(3,k) +","+arrayInput(4,k)+","+ "'"+arrayInput(5,k)+"')"
                  'response.write "</br>"+sql+"boom"
                  
                  cn.execute sql,numa 
              else
                  sql ="INSERT INTO je (sourceref, srseq, jemajor, jeminor, jesub1, jesub2, jeamount, jedesc) VALUES ("
                  sql = sql + cstr(srcnum) + ","+cstr(k+1)+ ","+ arrayInput(0,k) + "," + arrayInput(1,k)
                  sql = sql + "," + arrayInput(2,k) + "," + arrayInput(3,k) +","+arrayInput(4,k)+","+"' '" +")"
                  'response.write "</br>"+sql    
                  
                  cn.execute sql,numa 

              end if
              if numa=1 then
                  response.write "<P>Account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k) + " successfully inserted.</br>"
              else
                  response.write  "<P>Account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k) + " failed to inserted."
              end if

              next
          else
              response.write what_error
      end if
          
      'Updating glmaster

      if err_count = 0 then
          response.write "</br>-----------Updating account balances-----------</br>"
          for k = 0 to Request.Form("numvalid")-1
              
              newBal = cdbl(arrayInput(4,k))
              balances(k) = cdbl(balances(k)) + newBal
              sql ="UPDATE glmaster set balance = "+ cstr(balances(k)) +" where major=" +arrayInput(0,k)
              sql=sql+" AND minor="+arrayInput(1,k)
              sql=sql+" AND sub1="+arrayInput(2,k)
              sql=sql+" AND sub2="+arrayInput(3,k)
              'response.write "</br>"+sql
                  
              cn.execute sql,numa 
                  
              if numa=1 then
                  response.write "<P>Balance of account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k) + " successfully updated.</br>"
              else
                  err_count = err_count+1    
                  response.write  "<P>Balance of account major " +arrayInput(0,k)+ " minor "+arrayInput(1,k)+ " sub1 "+ arrayInput(2,k) +" sub2 " +arrayInput(3,k) + " failed to update.</br>"
              end if

          next
      else
          
      end if
      'Commit
      if err_count =0 then
          'Display accounts
          cn.CommitTrans
          response.write "<p><table border='1'><tr><td colspan='6' bgcolor='#aaaaaa' align='center'><b>Account Balances</td></tr><tr>"
          response.write "<td align='center'>Major</td>"
          response.write "<td align='center'>Minor</td>"
          response.write "<td align='center'>Sub1</td>"
          response.write "<td align='center'>Sub2</td>"
          response.write "<td align='center'>Account Description</td>"
          response.write "<td align='center'>Account<br>Balance</td></tr>"

          c=0

          for k=0 to request.form("numvalid")-1
              sql="SELECT * FROM glmaster where major=" +arrayInput(0,k)
              sql=sql+" AND minor="+arrayInput(1,k)
              sql=sql+" AND sub1="+arrayInput(2,k)
              sql=sql+" AND sub2="+arrayInput(3,k)
          
              set rs=Server.CreateObject("ADODB.Recordset")
              rs.open sql,"DSN=gl12345;UID=gl12345;PWD=QWERTYYUIOP;"
              while not rs.eof
                  response.write "<tr><td align='center'>"
                  response.write cstr(rs("major"))+"</td><td align='right'>"
                  response.write cstr(rs("minor"))+"</td><td align='right'>"
                  response.write cstr(rs("sub1"))+"</td><td align='right'>"
                  response.write cstr(rs("sub2"))+"</td><td align='right'>"
                  response.write cstr(rs("acctdesc"))+"</td><td align='right'>"
                  response.write formatnumber(rs("balance"),2)+"</td></tr>"
                  c=c+1
              rs.movenext

              wend
              rs.close
              set rs=nothing

          next

          'Display je
          response.write "<tr><td colspan='8' bgcolor='#aaaaaa' align='center'><b>Journal Entries</td></tr><tr>"
          response.write "<td align='center'>Source Reference Number</td>"
          response.write "<td align='center'>Sequence</td>"
          response.write "<td align='center'>Major</td>"
          response.write "<td align='center'>Minor</td>"
          response.write "<td align='center'>Sub1</td>"
          response.write "<td align='center'>Sub2</td>"
          response.write "<td align='center'>Transaction Description</td>"
          response.write "<td align='center'>Transaction Amount</td></tr>"


          for k=0 to request.form("numvalid")-1
              sql="SELECT * FROM je where sourceref=" +srcnum
              sql=sql+" AND srseq="+cstr(k+1)
          
              set rs=Server.CreateObject("ADODB.Recordset")
              rs.open sql,"DSN=gl12345;UID=gl12345;PWD=QWERTYYUIOP;"
              while not rs.eof
                  response.write "<tr><td align='center'>"
                  response.write cstr(rs("sourceref"))+"</td><td align='right'>"
                  response.write cstr(rs("srseq"))+"</td><td align='right'>"
                  response.write cstr(rs("jemajor"))+"</td><td align='right'>"
                  response.write cstr(rs("jeminor"))+"</td><td align='right'>"
                  response.write cstr(rs("jesub1"))+"</td><td align='right'>"
                  response.write cstr(rs("jesub2"))+"</td><td align='right'>"
                  response.write cstr(rs("jedesc"))+"</td><td align='right'>"
                  response.write formatnumber(rs("jeamount"),2)+"</td></tr>"
                  c=c+1
              rs.movenext

              wend
              rs.close
              set rs=nothing

          next

          response.write "</table><p><b>"+cstr(c)+" matching records touched. Data committed"
      else
          cn.RollbackTrans
          response.write "</br>Data not committed"

      end if
          cn.close
          set cn=nothing
      %>













            </div>
          </div>
      </div>

  </body>
</html>