
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/pads/model-pads.asp"-->

<%

Dim Conn
Dim Cmd
Dim listaPads
Dim combo
Dim idPad
Dim urlAtual
Dim urlOrigem

urlAtual = Request.ServerVariables("URL")
urlOrigem = Request.ServerVariables("HTTP_REFERER")

KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Request.Cookies("email")
email = py_decrypt(email, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta


consultaPads


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
