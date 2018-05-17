<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/hex_sha1_js.asp"-->
<%
 
Dim strName
Dim strInput, strOutput
 
strName = "Cali"
 
strInput = strName 
strOutput = hex_sha1(strInput)
 
Response.Write(strOutput)

'Tem que achar um jeito de mandar toda e qualquer senha encriptada para o ASP do back-end
%>