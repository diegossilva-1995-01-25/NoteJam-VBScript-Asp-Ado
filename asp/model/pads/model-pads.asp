
<%
	
	'Adicionar uma nova pasta
	Sub criarTab
	
		Dim comandoSqlPads

		nome = Server.htmlEncode(Request.QueryString("name"))
		regUsr = resgatarUsuario
		regUsr = Split(regUsr, ";")
		id = CInt(regUsr(0))
		email = regUsr(1)
	
		If Len(nome) > 0  AND Len(nome) <= 100 Then
	
			comandoSqlPads = "INSERT INTO pads(name, user_id) VALUES (?, ?);"
			
			conectar
			
			Cmd.Parameters.append Cmd.createParameter("@name", adVarChar, adParamInput, 100)
			Cmd.Parameters("@name").value = nome
			Cmd.Parameters.append Cmd.createParameter("@user_id", adInteger, adParamInput, 4)
			Cmd.Parameters("@user_id").value = id
	
			Cmd.commandText = comandoSqlPads
			Cmd.Execute  , , adExecuteNoRecords
							
			finalizar
			
			Set comandoSqlPads = nothing
		
			Response.Cookies("mensagem") = "Pad was, successfully, created."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
				
		ElseIf Len(nome) <= 0 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be empty or null."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp")
				
		ElseIf Len(nome) > 25 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be too long. The maximum allowed length is 100 characters."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp")
				
		End If
	
	
	End Sub


	'Alterar o nome da pasta
	Sub alterarPad
		
		Dim comandoSqlPads
		Dim pastaId
		
		nome = Server.htmlEncode(Request.QueryString("name"))
		pastaId = padId
	
		If Len(nome) > 0  AND Len(nome) <= 100 Then
	
			comandoSqlPads = "UPDATE pads SET name = ? WHERE id = ?;"
			
			conectar
		
			Cmd.Parameters.append Cmd.createParameter("@name", adVarChar, adParamInput, 100)
			Cmd.Parameters("@name").value = nome
			Cmd.Parameters.append Cmd.createParameter("@id", adInteger, adParamInput, 4)
			Cmd.Parameters("@id").value = pastaId
	
			Cmd.commandText = comandoSqlPads
			Cmd.Execute  , , adExecuteNoRecords
		
			finalizar
			
			Set comandoSqlPads = nothing
		
			Response.Cookies("mensagem") = "Note was, successfully, altered."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
				
		ElseIf Len(nome) <= 0 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be empty or null."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp?id=" & pastaId & "&name=" & nome)
				
		ElseIf Len(nome) > 25 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be too long. The maximum allowed length is 100 characters."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp?id=" & pastaId & "&name=" & nome)
				
		End If
	
	
	End Sub


	'Deletar uma pad
	Sub excluirNota
		
		Dim comandoSqlPads
		
		comandoSqlPads = "DELETE FROM pads WHERE id = ?"
		
		conectar
		
		Cmd.Parameters.append Cmd.createParameter("@id", adInteger, adParamInput, 4)
		Cmd.Parameters("@id").value = id
	
		Cmd.commandText = comandoSqlPads
		Cmd.Execute  , , adExecuteNoRecords
	
		finalizar
	
		Response.Cookies("mensagem") = "Pad was, successfully, deleted."
		Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
			
	End Sub


	'Trazer todas as pastas de notas à tela
	Sub consultaPads
		
		Dim comandoSqlPads
		comandoSqlPads = "SELECT pads.id, pads.name FROM pads INNER JOIN users ON pads.user_id = users.id WHERE users.email = ?;"
		
		conectar
		
		Cmd.Parameters.append Cmd.createParameter("@email", adVarChar, adParamInput, 75)
		Cmd.Parameters("@email").value = email
	
		Cmd.CommandType = adCmdText
		Cmd.CommandText = comandoSqlPads
		Cmd.ActiveConnection = Conn
		Cmd.ActiveConnection.CursorLocation = adUseClient
		
		Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.CursorType = adOpenStatic
		rs.Open(Cmd.execute)
	
		If rs.BOF AND rs.EOF Then
		
			listaPads = "No Pad"
		
		Else
	
			Response.buffer = True
			
			Dim padsArr()
			Dim options()
			
			Redim padsArr(rs.RecordCount)
			Redim options(rs.RecordCount)
			
			Dim index
			index = 0
			Dim pad
			Dim idPad
			
			DO UNTIL rs.EOF
		
				Dim auxPad
				idPad = rs.Fields.Item(0).Value
				pad = rs.Fields.Item(1).Value
				auxPad = encodeParaURL(pad)
				
		
				'Vai dentro das tags da linha
				Dim ref, altRef, delRef
				ref = "href=""/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp?id=" & idPad & "&name=" & pad & """"
				altRef = "href=""/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp?id=" & idPad & "&name=" & pad & """"
				delRef = "href=""/diego.silva/notejam-vbs-asp-ado/html/delete-pad.asp?id=" & idPad & "&name=" & pad & """"
					
				'Cada linha
				padsArr(index) = "<li><a " & ref & ">" & pad & "</a>" & _
					"<a " & altRef & "><img class=""imgAlt"" src=""/diego.silva/notejam-vbs-asp-ado/html/svg/edit.svg"" alt=""""></a>" & _
					"<a " & delRef & "><img class=""imgDel"" src=""/diego.silva/notejam-vbs-asp-ado/html/svg/delete.svg"" alt=""""></a> </li>"
		
				'Cada item do combobox
				options(index) = "<option value=""" & idPad & """>" & pad & "</option>"
				
				index = index + 1
				
				rs.moveNext
		
			LOOP
			
			listaPads = padsArr
			combo = options
	
		End If
	
		finalizar
	
	End Sub
	
	
	Sub consultaUmaPad
		
		Dim comandoSqlPads
	
		entrar
	
	
		comandoSqlPads = "SELECT pads.id, pads.name FROM pads INNER JOIN users ON pads.user_id = users.id WHERE users.email = ? AND pads.name = ?;"
		Cmmd.Parameters.append Cmmd.createParameter("@email", adVarChar, adParamInput, 75)
		Cmmd.Parameters("@email").value = mail
		Cmmd.Parameters.append Cmmd.createParameter("@padName", adVarChar, adParamInput, 100)
		Cmmd.Parameters("@padName").value = umPad
	
		Cmmd.commandText = comandoSqlPads
		Dim recSet	
		Set recSet = Cmmd.Execute
	
		DO UNTIL recSet.EOF
		
			idPasta = recSet.Fields.Item(0).Value
			umPad = recSet.Fields.Item(1).Value
		
			recSet.moveNext
	
		LOOP
			
		sair
	
	End Sub

	
%>

<script language="JavaScript" runat="server">
	function encodeParaURL(umaString) {
		return escape(umaString); //Será que cabe um laço aqui para converter caractere por caractere de   à %2F
	}
</script>
