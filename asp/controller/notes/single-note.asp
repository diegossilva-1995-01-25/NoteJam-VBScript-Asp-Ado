
<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/model/notes/model-notes.asp"-->

<%

Dim Con
Dim Com

KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
email = Request.Cookies("email")
email = py_decrypt(email, KEYTOCRYPT) 'Joga o valor do cookie na variável e só depois decripta

consultarUmaNota


'Método padrão para conexão do banco, com nomes alterados para não provocar conflito
Sub plug
	
	set Con = Server.createObject("ADODB.Connection") 
	Con.open("DRIVER=SQLite3 ODBC Driver; Database=" & Server.mapPath("/writables/db/diego.silva.sqlite") & "; LongNames=0; Timeout=1000; NoTXN=0; SyncPragma=NORMAL; StepAPI=0;")
	
	Set Com = Server.createObject("ADODB.Command") 
	Com.activeConnection = Con
	
End Sub


'Método padrão para desconexão do banco, com nomes para não provocar conflito
Sub overAndOut

	Con.close

	Set Com = nothing
	Set Con = nothing

End Sub


%>

