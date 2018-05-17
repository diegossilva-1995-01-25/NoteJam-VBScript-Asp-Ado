
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/login/model-login.asp"-->
<!--#include virtual="/lib/crypto.asp"-->

<%
	
	Dim Conn
	Dim Cmd
	Dim pwd
	Dim msg

	'Chamada do método para ver se o usuário existe
	verificarExistencia
	
	Sub conectar
	
		set Conn = Server.createObject("ADODB.Connection") 
		Conn.open("DRIVER=SQLite3 ODBC Driver; Database=" & Server.mapPath("/writables/db/diego.silva.sqlite") & "; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;")
		
		Set Cmd = Server.createObject("ADODB.Command") 
		Cmd.activeConnection = Conn
		
	End Sub
	
	
	Sub finalizar
	
		Conn.close
	
		Set Cmd = nothing
		Set Conn = nothing
	
	End Sub
	
	
	Sub verificarExistencia
		
		conectar
		
		email = Request.QueryString("email")
		
		'Busca no banco, por um usuário
		Dim sql
		sql = "SELECT email FROM users WHERE email = ?;"
		Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
		Cmd.Parameters("@email").value = email
		
		Cmd.commandText = sql
		Set rs = Cmd.execute
		
		'Se existe, cria a mensagem positiva e gera senha
		If rs.BOF AND rs.EOF Then
		
			Response.Cookies("mensagem") = "This e-mail doesn't exists in database."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/forgot-password.asp")
			
			
		ElseIf StrComp(email, rs.Fields.Item(0).Value, 0) = 0 Then
			
			Response.Cookies("mensagem") = "A link to reset to a brand new password has been, successfully, sent to your e-mail."
			enviarEmail email, geradorLink(email)
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		
		Else
			
			Response.Cookies("mensagem") = "This e-mail doesn't exists in database."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/forgot-password.asp")
			
		End If
						
		finalizar
		
		
	End Sub
	
	
	'Aqui será gerada uma nova senha
	Function gerador
		
		Dim randomico
		Dim num
		Set gerador = nothing
		
		letras = "abcdefghijklmnopqrstuvwxyz1234567890"
		
		'TEM que por o randomize para que ele não exiba os mesmos valores sempre
		Randomize
		For i = 1 to 8
			num = CInt((Len(letras) * Rnd) + 1)
			randomico = randomico & Mid(letras, num, 1)
		Next
		
		'Retorno
		gerador = randomico
		
		
	End Function
	
	
	'Aqui será gerada uma encriptação do link
	Function geradorLink(email)
		
		'Gerar um bcrypt do e-mail
		Dim CLASSIFIED
		CLASSIFIED = "MY_SPECIAL_PHRASE_TO_CYPHER_AN_INCOMING_DATA"
		geradorLink = py_encrypt(email, CLASSIFIED)
		
	End Function
	
%>

