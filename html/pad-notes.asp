<%@
language = "VBScript"
codepage = 65001
lcid     = 1033
%>
<%
option explicit
Response.buffer = true
Response.charset = "UTF-8"
%>
<!DOCTYPE html>
<!--[if lt IE 7 ]><html class="ie ie6" lang="en"> <![endif]-->
<!--[if IE 7 ]><html class="ie ie7" lang="en"> <![endif]-->
<!--[if IE 8 ]><html class="ie ie8" lang="en"> <![endif]-->
<!--[if (gte IE 9)|!(IE)]><!-->
 <!--<![endif]-->
<html lang="en">
<head>
	
	<%
		Response.CacheControl = "no-store"
		Dim TIMEOUT
		Dim secs
		TIMEOUT = 5 'minutos
		If Session("user") = "" Then
			Response.redirect("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		End If
	%>
	
	<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/controller/pads/show-pads.asp"-->
	<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/controller/notes/show-notes.asp"-->
	<!--#include virtual="/lib/crypto.asp"-->
	<%
		
		Dim email
		Dim lista
		Dim opcoes
		Dim notes
		Dim order
		Dim links
		Dim tablePt1
		Dim tablePt2
		Dim nomePad
		Dim nReg
		Dim ref
		Dim padId
		Dim KEYTOCRYPT
		KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
		
		email = Session("user")
		Session.Timeout = TIMEOUT
		secs = CInt((TIMEOUT * 60) + 1)
		order = Request.QueryString("order")
				
		If StrComp(TypeName(listaPads), "String") = 0 Then
			lista = listaPads
			opcoes = combo
		Else
			lista = join(listaPads)
			opcoes = join(combo)
		End If
		
		If StrComp(TypeName(listaNotes), "String") = 0 Then
			notes = listaNotes
		Else
			notes = join(listaNotes)
		End If
	
		links = pageLinks
	
		nReg = nNotes
	
		pad = Request.QueryString("name")
		padId = Request.QueryString("id")
		pad = Replace(pad, "%20", " ")
		nomePad = pad
	
		tablePt1 = "<table class=""notes"" id=""tabela""> <thead> <tr> " & _
			"<th class=""note"">Note <a href=""/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp?id=" & padId & "&pad=" & pad & "&order=noteasc"" class=""sort_arrow"" >▲</a>" & _
			"<a href=""/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp?id=" & padId & "&pad=" & pad & "&order=notedesc"" class=""sort_arrow"" >▼</a></th>" & _
			"<th>Pad</th>" & _
			"<th class=""date"">Last modified <a href=""/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp?id=" & padId & "&pad=" & pad & "&order=dateasc"" class=""sort_arrow"" >▲</a>" & _
			"<a href=""/diego.silva/notejam-vbs-asp-ado/html/pad-notes.asp?id=" & padId & "&pad=" & pad & "&order=datedesc"" class=""sort_arrow"" >▼</a></th>" & _
			"</tr> </thead> <tbody>"
		tablePt2 = "</tbody> </table>"
	
		if nReg = 0 Then
			tablePt1 = ""
			tablePt2 = ""
		End If
	
	
	
		ref = "href=""/diego.silva/notejam-vbs-asp-ado/html/create-pad.asp?id=" & padId & "&name=" & Server.urlEncode(pad) & """"
	
	%>
	
	<!-- Basic Page Needs
  ================================================== -->
	<meta charset="utf-8">
	<title>Notejam: <%=nomePad%>(<%=nReg%>)</title>
	<meta name="description" content="">
	<meta name="author" content="Diego S. Silva">
	<link rel="icon" href="/diego.silva/assets/img/radioactive.ico">
	<script src="/diego.silva/notejam-vbs-asp-ado/javascript/moment.js"></script>
	<script src="/diego.silva/notejam-vbs-asp-ado/javascript/moment-with-locales.js"></script>
	<meta http-equiv="refresh" content="<%=secs%>">

	<!-- Mobile Specific Metas
  ================================================== -->
	<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">

	<!-- CSS
  ================================================== -->
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/base.min.css">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/skeleton.min.css">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/layout.css">
	<link rel="stylesheet" href="/diego.silva/notejam-vbs-asp-ado/html/css/style.css">

	<!--[if lt IE 9]>
		<script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
	<![endif]-->
</head>
<body>
  <div class="container">
    <div class="sixteen columns">
      <div class="sign-in-out-block">
        <p>Welcome, <%=email%>:&nbsp; <a href="account-settings.asp">Account settings</a>&nbsp;&nbsp;&nbsp;<a href="signin.asp">Sign out</a></p>
      </div>
    </div>
    <div class="sixteen columns">
      <h1 class="bold-header"><a href="index.asp" class="header">note<span class="jam">jam: </span></a> <span><%=nomePad%> (<%=nReg%>)</span></h1>
    </div>
    <div class="three columns">
      <h4 id="logo">My pads</h4>
      <nav>
      <ul>
        <%=lista%>
      </ul>
      <hr />
      <a href="create-pad.asp">New pad</a>
      </nav>
    </div>
    <div class="thirteen columns content-area">
      <!--<div class="alert-area">-->
        <!--<div class="alert alert-success">Note is sucessfully saved</div>-->
      <!--</div>-->
       <%=tablePt1%>
         <%=notes%>
       <%=tablePt2%>
      <a class="newNote" href="/diego.silva/notejam-vbs-asp-ado/html/create.asp">New note</a>
      <a <%=ref%>>Pad Settings</a>
      <div class="pagination">
       <%=links%>
      </div>
    </div>
    <hr class="footer" />
    <div class="footer">
      <div>Notejam: <strong>VBScript + ASP + ADO</strong> application</div>
      <div>Version by <a href="https://github.com/diegossilva-1995-01-25">Diego S. Silva</a></div>
      <div><a href="https://github.com/komarserjio/notejam">Github</a>, <a href="https://twitter.com/komarserjio">Twitter</a>, based in the app created by <a href="https://github.com/komarserjio/">Serhii Komar</a></div>
    </div>
  </div><!-- container -->
  <script language="javascript">
	
	function dataCal() {
	
		var tabela = document.getElementById("tabela");
		
		if(!(tabela === null)) {
		
			for(var i = 1; i < tabela.rows.length; i++) {
			
				tabela.rows[i].cells[2].innerHTML = moment(String(tabela.rows[i].cells[2].innerHTML), "MM/DD/YYYY hh:mm:ss a").calendar();
			
			}
			
		}
		
	}
	
	dataCal();
	
  </script>
</body>

</html>


