
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/login/model-login.asp"-->
<!--#include virtual="/lib/crypto.asp"-->
<%

Dim Conn
Dim Cmd
Dim email
Dim pwd
Dim conf
Dim current
Dim urlOrigem

urlOrigem = Request.ServerVariables("HTTP_REFERER")
KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Request.Cookies("email")
email = py_decrypt(email, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta
current = Request.QueryString("current-password")
pwd = Request.QueryString("new-password")
conf = Request.QueryString("confirm-new-password")

If Len(current) > 0 Then
	alterarSenha
Else
	alterarSenhaForgotten email, pwd, conf	
End If


'Pega a senha atual do usuário no banco de dados
Function retornarSenha

	conectar
	
	comandoSql = "SELECT password FROM users WHERE email = ?;"
	Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
	Cmd.Parameters("@email").value = email
	
	Cmd.commandText = comandoSql
	Set rs = Cmd.execute
	
	retornarSenha = rs.Fields.Item(0).Value
	
	finalizar
	
End Function


'Método padrão para conexão do banco
Sub conectar

	set Conn = Server.createObject("ADODB.Connection") 
	Conn.open("DRIVER=SQLite3 ODBC Driver; Database=" & Server.mapPath("/writables/db/diego.silva.sqlite") & "; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;")
	
	Set Cmd = Server.createObject("ADODB.Command") 
	Cmd.activeConnection = Conn

End Sub


'Método padrão para desconexão do banco
Sub finalizar
	
	Conn.close
	
	Set Cmd = nothing
	Set Conn = nothing
	
End Sub

%>

