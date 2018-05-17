<!--#include virtual="/lib/crypto.asp"-->
<%

	Dim Conn
	Dim Cmd
	
	'Enviará o e-mail com a nova senha
	Sub enviarEmail (email, link)
		
		eml = "diego00023@gmail.com"
		pwd = "marioluigi"
		
		Set cdoMessage = CreateObject("CDO.Message")
		
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
		'Configuramos o email a ser usado para o envio da mensagem
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = eml

		'Inserimos a senha do email usado na linha de cima
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = pwd

		'Configuramos o tempo da tentativa de conexão
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60

		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'Name or IP of remote SMTP server
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
		'Server port
		cdoMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
	
		cdoMessage.From = eml
		cdoMessage.ReplyTo = eml
		cdoMessage.BodyPart.Charset = "utf-8"
		cdoMessage.Subject = "Your New Password to NOTEJAM"
		cdoMessage.TextBody = "Pleased user " & email & ":" & vbCrLf & vbCrLf & "Due security matters we've updated our system to make your accesses" & _
			" more safe. Said that so, we're sending you this link to you reset your password to continue using our NoteJam:" & vbCrLf & _
			"https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/reset-password.asp?e=" & link & vbCrLf & vbCrLf & _	
			"Respectfully" & vbCrLf & vbCrLf & "Diego S Silva"
		cdoMessage.To = email
	
		cdoMessage.Configuration.Fields.Update
	
		cdoMessage.Send
		
		Set cdoMessage = nothing
		
	End Sub
	
	
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
	
	
	'Aqui será gerada uma encriptação do link
	Function geradorLink(email)
		
		'Gerar um bcrypt do e-mail
		CLASSIFIED = "MY_SPECIAL_PHRASE_TO_CYPHER_AN_INCOMING_DATA"
		geradorLink = py_encrypt(email, CLASSIFIED)
		
	End Function
	

	Sub buscarEndEmail
	
		Dim mailAdd
		Dim l
	
		conectar
		
		comandoSql = "SELECT email FROM users;"
		Cmd.commandText = comandoSql
		Cmd.CommandType = adCmdText
		Cmd.CommandText = comandoSql
		Cmd.ActiveConnection = Conn
		Cmd.ActiveConnection.CursorLocation = adUseClient
		
		Set rs = Cmd.execute
		
		Do Until rs.EOF
			
			mailAdd = rs.Fields.Item(0).Value
			l = geradorLink(mailAdd)
			
			enviarEmail mailAdd, l
			
			rs.moveNext
			
		Loop
		
		
		finalizar
		
		Response.write("Over.")
		
	End Sub
	
	
	buscarEndEmail
	
	
%>