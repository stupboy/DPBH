
<%
'-数据库连接文件-
set conn=server.CreateObject("adodb.connection")
'-".""为服务器地址、ST为连接数据库名称、sa为数据库用户名、PWD为数据库密码-
ConnStr="server=.;driver={sql server};database=ST;uid=sa;pwd=!@#$%asdfg"
conn.Open connstr
'-如果连接出错则报错-
If Err Then
  err.Clear
  Set Conn = Nothing
  Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
  Response.End
End If
'-Oracle数据库-
If Not IsObject(Conn) Then
Set Conn1 = Server.CreateObject("ADODB.Connection")
myDSN = "Provider=OraOLEDB.Oracle;Data Source=242;User ID=neands3;PASSWORD=abc123;Persist Security Info=True"
Conn1.Open myDSN
End If

If Err Then
  err.Clear
  Set Conn1 = Nothing
  Response.Write "数据库连接出错，请检查Conn.asp文件中的数据库参数设置。"
  Response.End
End If
%>