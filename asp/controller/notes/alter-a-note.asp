
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/notes/model-notes.asp"-->
<!--#include virtual="/lib/crypto.asp"-->
<%

Dim Conn
Dim Cmd

urlOrigem = Request.ServerVariables("HTTP_REFERER")
urlOrigem = Split(urlOrigem, "?id=")
segundoSplit = Split(urlOrigem(1), "&")
id = segundoSplit(0)
KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Request.Cookies("email")
email = py_decrypt(email, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta

alterarNota


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

