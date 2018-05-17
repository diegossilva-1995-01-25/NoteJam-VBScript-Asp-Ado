
<!--#include virtual="/lib/crypto.asp"-->
<!--#include virtual="/lib/packages.asp"-->


<%
	Dim comandoSqlLogin
	Dim KEYTOCRYPT
	KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
	
	'As tabelas a serem criadas no sistema serão criadas aqui, se não existirem
	Sub criarTabelas
	
		comandoSqlLogin = "CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, email VARCHAR(75) NOT NULL, password VARCHAR(128) NOT NULL);" & _
		"CREATE TABLE IF NOT EXISTS pads (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, name VARCHAR(100) NOT NULL, user_id INTEGER NOT NULL REFERENCES users(id));" & _
		"CREATE TABLE IF NOT EXISTS notes (id INTEGER PRIMARY KEY AUTOINCREMENT NOT NULL, pad_id INTEGER REFERENCES pads(id), " & _
		"user_id INTEGER NOT NULL REFERENCES users(id), name VARCHAR(100) NOT NULL, text text NOT NULL, created_at DATETIME NOT NULL, updated_at DATETIME NOT NULL);"
		
		conectar
		
		Cmd.commandText = comandoSqlLogin
		Cmd.Execute  , , adExecuteNoRecords
				
		finalizar
		
		Set comandoSqlLogin = nothing
	
	End Sub


	'O novo usuário será criado
	Sub criarUsuario
		
		email = Server.htmlEncode(Request.QueryString("email"))
		senha = Server.htmlEncode(Request.QueryString("password"))
		confirma = Server.htmlEncode(Request.QueryString("confirm-password"))
		
		conectar
	
		If StrComp(senha, confirma, 0) = 0 AND Len(senha) >= 8 AND _
			Len(email) > 0 AND InStr(email, "@") > 0 Then
			
			senha = py_digest(senha) 'Gerar o hash
	
			comandoSqlLogin = "INSERT INTO users(email, password) VALUES (?, ?);"
			Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
			Cmd.Parameters("@email").value = email
			Cmd.Parameters.append Cmd.createParameter("@password", adVarChar, adParamInput, 128)
			Cmd.Parameters("@password").value = senha
	
			Cmd.commandText = comandoSqlLogin
			Cmd.Execute  , , adExecuteNoRecords
		
			finalizar
			
			Session("user") = email
			email = py_encrypt(email, KEYTOCRYPT)
			Response.Cookies("email") = email 'Encripta e joga na variável, depois joga no cookie
		
			Response.Cookies("mensagem") = "Welcome to NoteJam. Take note of everything you want, my friend."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
	
		ElseIf Len(email) <= 0 Then
		
			finalizar
			Response.Cookies("mensagem") = "E-mail is required."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signup.asp")
		
		ElseIf InStr(email, "@") = 0 Then
		
			finalizar
			Response.Cookies("mensagem") = "E-mail is not valid."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signup.asp")
		
		ElseIf StrComp(senha, confirma, 0) <> 0 Then
		
			finalizar
			Response.Cookies("mensagem") = "Password doesnt match with the confirmation."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signup.asp")
		
		ElseIf Len(senha) < 8 Then
	
			finalizar
			Response.Cookies("mensagem") = "Passwords must not have less then 8 characters."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signup.asp")
	
		End If
	
	End Sub
	
	
	Sub validarUsuario
				
		email = Server.htmlEncode(Request.QueryString("email"))
		senha = Server.htmlEncode(Request.QueryString("password"))
		
		conectar
	
		If Len(email) <= 0 Then
		
			finalizar
			Response.Cookies("mensagem") = "E-mail is required."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		
		ElseIf Len(senha) <= 0 Then
		
			finalizar
			Response.Cookies("mensagem") = "Password is required."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		
		Else
	
			comandoSqlLogin = "SELECT password FROM users WHERE email = ?;"
			Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
			Cmd.Parameters("@email").value = email
	
			Cmd.commandText = comandoSqlLogin
			Set rs = Cmd.execute
		
			If rs.BOF AND rs.EOF Then
			
				Response.Cookies("mensagem") = "Your e-mail is not signed in."
				Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
			
			Else
		
				senhaRetornada = rs.Fields.Item(0).Value
				
				finalizar
	
				If py_digest_chk(senha, senhaRetornada) Then 'Se os hashes (hash do banco e o hash gerado aqui) batem
		
					Session("user") = email
					email = py_encrypt(email, KEYTOCRYPT)
					Response.Cookies("email") = email 'Encripta e joga na variável, depois joga no cookie
					
					Response.Cookies("mensagem") = "Welcome back. Good thing you came."
					Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
							
				Else
		
					Response.Cookies("mensagem") = "Your password is wrong."
					Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
							
				End If
			
			End If
			
		
		End If
	
	End Sub


	' Só use para realizar testes com banco de dados
	Sub darDrop

		conectar
	
		'comandoSql = "DROP TABLE IF EXISTS notes; DROP TABLE IF EXISTS pads; DROP TABLE IF EXISTS users;"
		'comandoSql = "DELETE FROM users WHERE users.email = '';"
	
		Cmd.commandText = comandoSqlLogin
		Cmd.Execute  , , adExecuteNoRecords
	
		finalizar
	
		Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/signin.asp")

	End Sub
	
	
	'Alterará a senha no banco de dados
	Sub alterarSenhaForgotten (email, pwd, conf)
	
		If StrComp(pwd, conf, 0) <> 0 Then
			
			Response.Cookies("mensagem") = "The password doesnt match with its confirmation."
			Response.redirect(urlOrigem)
			
		Else
			
			pwd = py_digest(pwd) 'Gerar o hash
			
			conectar
						
			comandoSqlLogin = "UPDATE users SET password = ? WHERE email = ?;"
			Cmd.Parameters.append Cmd.createParameter("@password", adVarChar, adParamInput, 128)
			Cmd.Parameters("@password").value = pwd
			Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
			Cmd.Parameters("@email").value = email
		
			Cmd.commandText = comandoSqlLogin	
			Cmd.Execute  , , adExecuteNoRecords
		
			finalizar
		
			Session("user") = email
			email = py_encrypt(email, KEYTOCRYPT)
			Response.Cookies("email") = email 'Encripta e joga na variável, depois joga no cookie
			Response.Cookies("mensagem") = "Welcome. *sigh* Good thing you still with us, thought we would not see each other anymore."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
			
		End If
		
	End Sub
	
	
	'Enviará o e-mail com a nova senha
	Sub enviarEmail (email, link)
		
		eml = "diego00023@gmail.com"
		pwd = "marioluigi"
		
		Dim cdomensagem		
		Set cdomensagem = CreateObject("CDO.Message")
		
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
		'Configuramos o email a ser usado para o envio da mensagem
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = eml

		'Inserimos a senha do email usado na linha de cima
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pwd

		'Configuramos o tempo da tentativa de conexão
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'Name or IP of remote SMTP server
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
		'Server port
		cdomensagem.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	
		cdomensagem.From = eml
		cdomensagem.ReplyTo = eml
		cdomensagem.BodyPart.Charset = "utf-8"
		cdomensagem.Subject = "Your New Password to NOTEJAM"
		cdomensagem.TextBody = "Pleased user " & email & ":" & vbCrLf & vbCrLf & "That's the link to create your brand new password: " & _
			"https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/reset-password.asp?e=" & link & vbCrLf & vbCrLf & _	
			"Respectfully" & vbCrLf & vbCrLf & "Diego S Silva"
		cdomensagem.To = email
	
		cdomensagem.Configuration.Fields.Update
	
		cdomensagem.Send
		
		Set cdomensagem = nothing
		
	End Sub
	
	
	'Altera a senha se a atual estiver correta e a senha confere com a confirmação e se for maior que 8 caracteres
	Sub alterarSenha

		senhaAtual = Server.htmlEncode(Request.QueryString("current-password"))
		senhaRetornada = retornarSenha
	
	
		If py_digest_chk(senhaAtual, senhaRetornada) Then
	
	
			novaSenha = Server.htmlEncode(Request.QueryString("new-password"))
			confirmaSenha = Server.htmlEncode(Request.QueryString("confirm-new-password"))
		
			If StrComp(novaSenha, confirmaSenha, 0) = 0 AND Len(novaSenha) >= 8 Then
											
				novaSenha = py_digest(novaSenha) 'Gerar o hash
				
				conectar
		
				comandoSqlLogin = "UPDATE users SET password = ? WHERE email = ?;"
				Cmd.Parameters.append Cmd.createParameter("@password", adVarChar, adParamInput, 128)
				Cmd.Parameters("@password").value = novaSenha
				Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
				Cmd.Parameters("@email").value = email
		
				Cmd.commandText = comandoSqlLogin	
				Cmd.Execute  , , adExecuteNoRecords
		
				finalizar
			
				Response.Cookies("mensagem") = "Your password has been, successfully, altered."
				Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
			
			ElseIf Len(novaSenha) < 8 Then
			
				Response.Cookies("mensagem") = "Your password is too short. The minimum length allowed is 8 characters."
				Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/account-settings.asp")
			
			ElseIf StrComp(novaSenha, confirmaSenha, 0) <> 0 Then
			
				Response.Cookies("mensagem") = "Your password doesnt match with its confirmation."
				Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/account-settings.asp")
			
			End If
	
	
		Else
		
			Response.Cookies("mensagem") = "Your password is wrong."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/account-settings.asp")
		
		End If
	
	
	End Sub
	
	
	Sub encriptarSenhasExistentes
						
		comandoSqlLogin = "SELECT id, password FROM users;"
		
		Dim ids()
		Dim pwds()
						
		cont = 0
		
		conectar
		
		Cmd.commandText = comandoSqlLogin
		Cmd.CommandType = adCmdText
		Cmd.CommandText = comandoSqlLogin
		Cmd.ActiveConnection = Conn
		Cmd.ActiveConnection.CursorLocation = adUseClient
		Set rs = Cmd.execute
		
		n = rs.RecordCount
		Redim ids(n)
		Redim pwds(n)
		
		Do Until rs.EOF
			
			id = rs.Fields.Item(0).Value
			pwd = py_digest(rs.Fields.Item(1).Value) 'Gerar o hash
			
			ids(cont) = id
			pwds(cont) = pwd
			
			cont = cont + 1
			
		Loop
		
		finalizar
		
		For cont = 0 To n-1 Step 1
								
			comandoSql2 = "UPDATE users SET password = ? WHERE id = ?;"
			
			conectar
			
			Cmd.Parameters.append Cmd.createParameter("@password", adVarChar, adParamInput, 128)
			Cmd.Parameters("@password").value = pwds(cont)
			Cmd.Parameters.append Cmd.createParameter("@id", adVarChar, adParamInput, 4)
			Cmd.Parameters("@id").value = ids(cont)
		
			Cmd.commandText = comandoSql2	
			Cmd.Execute  , , adExecuteNoRecords
			
			finalizar
			
		Next
		
	End Sub
	
	
%>

