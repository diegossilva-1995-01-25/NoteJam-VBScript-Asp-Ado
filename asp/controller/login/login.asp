
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/login/model-login.asp"-->
<%
	
Dim Conn
Dim Cmd
Dim email
Dim urlOrigem
Dim urlPura

email = Server.htmlEncode(Request.QueryString("email"))

Set urlOrigem = Request.ServerVariables("HTTP_REFERER")

urlOrigem = Split(urlOrigem, "?")
urlPura = urlOrigem(0)


'Se a tela for para login, valida o usuário, senão usa os métodos de criação
If StrComp("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp", urlPura, 0) = 0 Then
	
	validarUsuario
				
ElseIf StrComp("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signup.asp", urlPura, 0) = 0 Then

	criarTabelas
	if verificarEmailCadastravel Then
		criarUsuario
	Else
		Response.Cookies("mensagem") = "This e-mail already exists in database."
		Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signup.asp")
	End If
			
End If


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


Function verificarEmailCadastravel
	
	Dim retorno
	retorno = false
	
	Dim sql
	email = Server.htmlEncode(Request.QueryString("email"))
	
	conectar
	
	'Busca no banco, por um usuário
	sql = "SELECT email FROM users WHERE email = ?;"
	Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
	Cmd.Parameters("@email").value = email
	
	Cmd.commandText = sql
	Set rs = Cmd.execute
	
	'Se existe, cria a mensagem positiva e gera senha
	If rs.BOF AND rs.EOF Then
	
		retorno = true
		
	End If
					
	finalizar
	
	verificarEmailCadastravel = retorno
	
End Function

	
%>
