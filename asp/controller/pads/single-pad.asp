
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/pads/model-pads.asp"-->

<%

Dim Connec
Dim Cmmd
Dim onePad
Dim idPasta
Dim umPad
Dim param
Dim mail

'urlAtual = Request.ServerVariables("QUERY_STRING")
umPad = Server.htmlEncode(Request.QueryString("pad"))
KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
mail = Request.Cookies("email")
mail = py_decrypt(mail, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta
'Response.write(umPad)

consultaUmaPad


'Método padrão para conexão do banco
Sub entrar
	
	set Connec = Server.createObject("ADODB.Connection") 
	Connec.open("DRIVER=SQLite3 ODBC Driver; Database=" & Server.mapPath("/writables/db/diego.silva.sqlite") & "; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;")
	
	Set Cmmd = Server.createObject("ADODB.Command") 
	Cmmd.activeConnection = Connec
	
End Sub


'Método padrão para desconexão do banco
Sub sair

	Connec.close

	Set Connec = nothing
	Set Conn = nothing

End Sub


%>

