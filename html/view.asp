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
	<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/controller/notes/single-note.asp"-->
	<!--#include virtual="/lib/packages.asp"-->
	<!--#include virtual="/lib/crypto.asp"-->
	<%
		
		Dim email
		Dim pad
		Dim lista
		Dim cod
		Dim title
		Dim text
		Dim data
		Dim paramet
		Dim paramet2
		Dim titulo
		Dim corpo
		Dim item
		Dim id
		Dim dataMod
		Dim KEYTOCRYPT
		KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
		
		email = Session("user")
		Session.Timeout = TIMEOUT
		secs = CInt((TIMEOUT * 60) + 10)
		pad = Server.htmlEncode(Request.QueryString("pad"))
	
		If StrComp(TypeName(listaPads), "String") = 0 Then
			lista = listaPads
		Else
			lista = join(listaPads)
		End If
		
		Dim auxTitle
		Dim auxText
	
		cod = id
		auxTitle = titulo
		auxText = encodeParaURL(corpo)
		title = "## " & titulo
		title = MarkDownIt.render(title)
		text = corpo
		text = MarkDownIt.render(text)
		data = dataMod
	
		paramet = "?id=" & cod & "&name=" & Server.urlEncode(auxTitle)
		paramet2 = "?id=" & cod & "&title=" & Server.urlEncode(auxTitle)
	%>
	
	<!-- Basic Page Needs
  ================================================== -->
	<meta charset="utf-8">
	<title>Notejam: <%=auxTitle%></title>
	<meta name="description" content="Tela para visualizar o conteúdo de uma nota selecionada">
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
      <h1 class="bold-header"><a href="index.asp" class="header">note<span class="jam">jam:</span></a> <span> <%=auxTitle%></span></h1>
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
      <p id="hidden" class="hidden-text">Last edited at <%=data%></p>
      <div class="note">
        <strong><%=title%></strong>
        <%=text%>
      </div>
      <!-- Coloque o nome da nota como parâmetro destes dois botões -->
      <a class="edit" href="edit.asp<%=paramet%>"> <button type="button">Edit</button> </a>
      <a href="delete.asp<%=paramet2%>" class="delete-note">Delete it</a>
    </div>
    <hr class="footer" />
    <div class="footer">
      <div>Notejam: <strong>VBScript + ASP + ADO</strong> application</div>
      <div>Version by <a href="https://github.com/diegossilva-1995-01-25">Diego S. Silva</a></div>
      <div><a href="https://github.com/komarserjio/notejam">Github</a>, <a href="https://twitter.com/komarserjio">Twitter</a>, based in the app created by <a href="https://github.com/komarserjio/">Serhii Komar</a></div>
    </div>
  </div><!-- container -->
  <script language="JavaScript">
	function dataCal() {
		var umaData = "<%=dataMod%>";
		var comp = document.getElementById("hidden");
		comp.innerHTML = "Last edited at " + moment(String(umaData), "MM/DD/YYYY hh:mm:ss a").subtract(3, "hours").calendar();
	}
	
	dataCal();
  </script>
  
  <script language="JavaScript" runat="server">
	function encodeParaURL(umaString) {
		return escape(umaString); //Será que cabe um laço aqui para converter caractere por caractere de %20 à %2F
	}
  </script>
</body>
</html>



