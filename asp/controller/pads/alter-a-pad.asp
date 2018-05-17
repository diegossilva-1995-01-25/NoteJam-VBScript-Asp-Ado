
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/pads/model-pads.asp"-->
<!--#include virtual="/lib/crypto.asp"-->
<%

Dim Conn
Dim Cmd
Dim padId

urlOrigem = Request.ServerVariables("HTTP_REFERER")
urlOrigem = Split(urlOrigem, "?id=")
segundoSplit = Split(urlOrigem(1), "&")
padId = segundoSplit(0)
KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Request.Cookies("email")
email = py_decrypt(email, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta

alterarPad


'Resgatar o usuário do banco para ser usado como filtro do SQL
Function resgatarUsuario

	conectar
		
	comandoSql = "SELECT id, email FROM users WHERE email = ?;"
	Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
	Cmd.Parameters("@email").value = email
	
	Cmd.commandText = comandoSql
	Set rs = Cmd.execute
	
	aux = rs.Fields.Item(0).Value & ";" & rs.Fields.Item(1).Value
	
	finalizar
		
	
	retorno = aux
	
	resgatarUsuario = retorno
	
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

