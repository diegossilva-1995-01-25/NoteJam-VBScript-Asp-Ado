
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/notes/model-notes.asp"-->

<%

Dim Conexao
Dim Commando
Dim listaNotes
Dim cont
Dim cont2
Dim paginas
Dim pageLinks
Dim nNotes
Dim page
Dim pad
Dim padParaUrl
Dim urlAgora
Dim urlOriginal
Dim queries

cont = 1
cont2 = 1
nNotes = 0

urlOriginal = Request.ServerVariables("HTTP_REFERER")
queries = Request.ServerVariables("QUERY_STRING")
urlAgora = Request.ServerVariables("URL")

KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Session("user")
order = Server.htmlEncode(Request.QueryString("order"))
padParaUrl = ""
pad = ""
page = 1

'Separa o número de página
If InStr(queries, "&page=") > 0 Then
	page = Server.htmlEncode(Request.QueryString("page"))
End If

If InStr(queries, "pad=") > 0 Then
	pad = Server.htmlEncode(Request.QueryString("pad"))
	padParaUrl = "pad="
End If


consultaNotes


'Método padrão para conexão do banco, nomes alterados para não ter conflito
Sub conex
	
	set Conexao = Server.createObject("ADODB.Connection") 
	Conexao.open("DRIVER=SQLite3 ODBC Driver; Database=" & Server.mapPath("/writables/db/diego.silva.sqlite") & "; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;")
	
	Set Commando = Server.createObject("ADODB.Command") 
	Commando.activeConnection = Conexao
	
End Sub


'Método padrão para desconexão do banco, nomes alterados para não ter conflito
Sub fim

	Conexao.close

	Set Commando = nothing
	Set Conexao = nothing

End Sub


'Método que cria os links para as páginas de registros
Sub criarLinksPagina

	Response.buffer = True
	
	Dim ref
	
	cont = 1
	pageLinks = ""
			
	While cont <= paginas
		
		if Len(padParaUrl) <> 0 Then
			ref = "href=""" & urlAgora & "?" & padParaUrl & pad & "&order=" & order & "&page=" & cont & """"
		Else
			ref = "href=""" & urlAgora & "?order=" & order & "&page=" & cont & """"
		End If
		
		pageLinks = pageLinks & ("<a " & ref & ">" & cont & "</a> ")
		cont = cont + 1
	Wend
	
End Sub


%>

