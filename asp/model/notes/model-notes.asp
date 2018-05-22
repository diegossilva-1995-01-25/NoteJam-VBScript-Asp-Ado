
<script src="/diego.silva/notejam-vbs-asp-ado/javascript/moment.js"></script>
<script src="/diego.silva/notejam-vbs-asp-ado/javascript/moment-with-locales.js"></script>

<%
	Dim comandoSqlNotes
	
	'Inserir a nota no banco de dados
	Sub criarNote
	
		nome = Server.htmlEncode(Trim(Request.Form("name")))
		texto = Server.htmlEncode(Trim(Request.Form("text")))
		pad = Server.htmlEncode(Request.Form("list"))
	
		regUsr = resgatarUsuario
		regUsr = Split(regUsr, ";")
		id = CInt(regUsr(0))
		email = regUsr(1)
	
		If Len(nome) > 0  AND Len(nome) <= 100 Then
		
			Dim comandoSqlNotes
			
			conectar
		
			If pad <> 0 Then
		
				comandoSqlNotes = "INSERT INTO notes(pad_id, user_id, name, text, created_at, updated_at) " & _
					"VALUES (?, ?, ?, ?, DATETIME('NOW'), DATETIME('NOW'));"
				Cmd.Parameters.append Cmd.createParameter("@pad_id", adInteger, adParamInput, 4)
				Cmd.Parameters("@pad_id").value = pad
			
			Else
			
				comandoSqlNotes = "INSERT INTO notes(pad_id, user_id, name, text, created_at, updated_at) " & _
					"VALUES (NULL, ?, ?, ?, DATETIME('NOW'), DATETIME('NOW'));"
			
			End If
		
			Cmd.Parameters.append Cmd.createParameter("@user_id", adInteger, adParamInput, 4)
			Cmd.Parameters("@user_id").value = id
			Cmd.Parameters.append Cmd.createParameter("@name", adVarChar, adParamInput, 100)
			Cmd.Parameters("@name").value = nome
			Cmd.Parameters.append Cmd.createParameter("@text", adVarChar, adParamInput, 2147483647)
			Cmd.Parameters("@text").value = texto
		
			Cmd.commandText = comandoSqlNotes
			Cmd.Execute  , , adExecuteNoRecords
		
			finalizar
			
			Set comandoSqlNotes = nothing
		
			Response.Cookies("mensagem") = "Note was, successfully, created."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
					
		ElseIf Len(nome) <= 0 Then
	
			Response.Cookies("mensagem") = "Note name cannot be empty or null."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create.asp")
				
		ElseIf Len(nome) > 100 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be too long. The maximum allowed length is 100 characters."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/create.asp")
	
		End If
	
	End Sub
	
	
	Sub alterarNota
	
		nome = Server.htmlEncode(Trim(Request.Form("name")))
		texto = Server.htmlEncode(Trim(Request.Form("text")))
		pad = Server.htmlEncode(Request.Form("list"))
		'msg = Server.htmlEncode(Request.QueryString("msg"))
		
		If Len(nome) > 0  AND Len(nome) <= 100 Then
	
			conectar
		
			If pad <> 0 Then
				comandoSqlNotes = "UPDATE notes SET pad_id = ?, name = ?, text = ?, updated_at = DATETIME('NOW') WHERE id = ?;"
				Cmd.Parameters.append Cmd.createParameter("@pad_id", adInteger, adParamInput, 4)
				Cmd.Parameters("@pad_id").value = pad
			Else
				comandoSqlNotes = "UPDATE notes SET pad_id = NULL, name = ?, text = ?, updated_at = DATETIME('NOW') WHERE id = ?;"
			End If
			
			Cmd.Parameters.append Cmd.createParameter("@name", adVarChar, adParamInput, 100)
			Cmd.Parameters("@name").value = nome
			Cmd.Parameters.append Cmd.createParameter("@text", adVarChar, adParamInput, 2147483647)
			Cmd.Parameters("@text").value = texto
			Cmd.Parameters.append Cmd.createParameter("@id", adInteger, adParamInput, 4)
			Cmd.Parameters("@id").value = id
		
			Cmd.commandText = comandoSqlNotes
			Cmd.Execute  , , adExecuteNoRecords
		
			finalizar
			
			Set comandoSqlNotes = nothing
		
			Response.Cookies("mensagem") = "Note was, successfully, altered."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
				
		ElseIf Len(nome) <= 0 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be empty or null."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/edit.asp?id=" & id & "&name=" & nAnterior)
				
		ElseIf Len(nome) > 100 Then
	
			Response.Cookies("mensagem") = "Pad name cannot be too long. The maximum allowed length is 100 characters."
			Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/edit.asp?id=" & id & "&name=" & nAnterior)
	
		End If
	
	End Sub
	
	
	'Deletar uma nota
	Sub excluirNota
	
		comandoSqlNotes = "DELETE FROM notes WHERE id = ?"
		
		conectar
		
		Cmd.Parameters.append Cmd.createParameter("@id", adInteger, adParamInput, 4)
		Cmd.Parameters("@id").value = id
	
		Cmd.commandText = comandoSqlNotes
		Cmd.Execute  , , adExecuteNoRecords
	
		finalizar
	
		Response.Cookies("mensagem") = "Note was, successfully, deleted."
		Response.redirect("/diego.silva/notejam-vbs-asp-ado/html/index.asp")
	
	End Sub
	
	
	'Buscará uma só nota no banco de dados para exibir na tela seu conteúdo
	Sub consultarUmaNota
	
		Dim nome
		Dim pasta

		nome = Request.QueryString("name")
		pasta = Server.htmlEncode(Request.QueryString("pad"))
	
		comandoSqlNotes = "SELECT notes.id, notes.name, notes.text, pads.id, notes.updated_at FROM notes " & _
			"LEFT JOIN pads ON pads.id = notes.pad_id " & _
			"INNER JOIN users ON users.id = notes.user_id WHERE users.email = ? " & _ 
			"AND notes.name = ?;"
	
		plug
		
		Com.Parameters.append Com.createParameter("@email", adVarChar, adParamInput, 75)
		Com.Parameters("@email").value = email
		Com.Parameters.append Com.createParameter("@note", adVarChar, adParamInput, 100)
		Com.Parameters("@note").value = nome
	
		Com.commandText = comandoSqlNotes	
		Dim rs
		Set rs = Com.Execute
	
		id = rs.Fields.Item(0).Value
		titulo = rs.Fields.Item(1).Value
		corpo = rs.Fields.Item(2).Value
		item = rs.Fields.Item(3).Value
		dataMod = rs.Fields.Item(4).Value
		
		overAndOut
	
	End Sub
	
	
	'Traz todas as notas do usuário
	Sub consultaNotes
			
		conex
				
		If InStr(urlAgora, "/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp") > 0 Then
		
			comandoSqlNotes = "SELECT notes.name, pads.name, notes.updated_at, pads.id FROM notes " & _
				"LEFT JOIN pads ON pads.id = notes.pad_id " & _
				"INNER JOIN users ON users.id = notes.user_id WHERE users.email = ? AND pads.id = ?"
			comandoSqlNotes = completarSql(comandoSqlNotes)
			pad = Request.QueryString("id")
	
			Commando.CommandType = adCmdText
			Commando.CommandText = comandoSqlNotes
			Commando.ActiveConnection = Conexao
			Commando.ActiveConnection.CursorLocation = adUseClient
			Commando.Parameters.append(Commando.createParameter("@email", adVarChar, adParamInput, 75))
			Commando.Parameters("@email").value = email
			Commando.Parameters.append(Commando.createParameter("@name", adVarChar, adParamInput, 100))
			Commando.Parameters("@name").value = pad		
		
		Else
		
			comandoSqlNotes = "SELECT notes.name, pads.name, notes.updated_at, pads.id FROM notes " & _
				"LEFT JOIN pads ON pads.id = notes.pad_id " & _
				"INNER JOIN users ON users.id = notes.user_id WHERE users.email = ?"
			comandoSqlNotes = completarSql(comandoSqlNotes)
	
			Commando.CommandType = adCmdText
			Commando.CommandText = comandoSqlNotes
			Commando.ActiveConnection = Conexao
			Commando.ActiveConnection.CursorLocation = adUseClient
			Commando.Parameters.append(Commando.createParameter("@email", adVarChar, adParamInput, 75))
			Commando.Parameters("@email").value = email
		
		End If
	
	
		Dim rec
		Set rec = Server.CreateObject("ADODB.Recordset")
		rec.CursorType = adOpenStatic
		rec.Open(Commando.execute)
	
		If rec.BOF AND rec.EOF Then
		
			pageLinks = ""
			listaNotes = "<p class=""first"">Create your first note</p><br>"
		
		Else
	
			Response.buffer = True
			
			rec.PageSize = 10
			rec.AbsolutePage = page
	
			paginas = rec.PageCount
	
			'Para ver quantos registros temos
			nNotes = rec.RecordCount
			
			Dim linhas(10)
			Dim nome
			Dim padName
			Dim modif
			Dim ref
			Dim ref2
	
			DO UNTIL rec.EOF Or cont > rec.PageSize '(Para exibir o número de notas daquela página)
				
				nome = rec.Fields.Item(0).Value
				padName = rec.Fields.Item(1).Value
				modif = rec.Fields.Item(2).Value
				pad = rec.Fields.Item(3).Value
		
				'Para dentro da tag
				ref = "href=""/diego.silva/notejam-vbs-asp-ado/html/view.asp?name=" & Server.urlEncode(nome) & "&pad=" & Server.urlEncode("" & padName) & """"
				ref2 = "href=""/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp?id=" & Server.urlEncode("" & pad) & "&name=" & Server.urlEncode("" & padName) & """"
		
				'Cada linha da tabela
				linhas(cont - 1) = "<tr> <td> <a " & ref & ">" & nome & "</a> </td> <td> <a " & ref2 & ">" & padName & "</a> </td> " & _
					"<td>" & modif & "</td> </tr> "
		
				'Se pad for null dizer que o nome é NO PAD
				If isNull(padName) Then
					padName = "NO PAD"
					
					'Cada linha da tabela
					linhas(cont - 1) = "<tr> <td> <a " & ref & ">" & nome & "</a> </td> <td>" & padName & "</td> " & _
						"<td>" & modif & "</td> </tr> "
					
				End If
					
				rec.moveNext
				cont = cont + 1
			
		
			LOOP
			
			listaNotes = linhas
	
			criarLinksPagina
		
		End If
			
		fim
	
	End Sub
	
	
	Function completarSql(comandoSql)

		'Para quando as setas de ordenação forem acionadas
		If StrComp(order, "noteasc", 0) = 0 Then
			comandoSql = comandoSql & " ORDER BY notes.name ASC;"
		ElseIf StrComp(order, "notedesc", 0) = 0 Then
			comandoSql = comandoSql & " ORDER BY notes.name DESC;"
		ElseIf StrComp(order, "dateasc", 0) = 0 Then
			comandoSql = comandoSql & " ORDER BY notes.updated_at ASC;"
		ElseIf StrComp(order, "datedesc", 0) = 0 Then
			comandoSql = comandoSql & " ORDER BY notes.updated_at DESC;"
		Else
			comandoSql = comandoSql & " ORDER BY notes.updated_at DESC;"
		End If
	
		completarSql = comandoSql

	End Function


%>

