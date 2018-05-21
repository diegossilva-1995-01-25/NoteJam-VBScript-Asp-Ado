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

<html lang="en"> <!--<![endif]-->
<head>
	
	<%
		Response.CacheControl = "no-store"
		Dim TIMEOUT
		Dim secs
		TIMEOUT = 120 'minutos
		If Session("user") = "" Then
			Response.redirect("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		End If
	%>
	
	<!--#include virtual="/diego.silva/notejam-vbs-asp-ado/asp/controller/pads/show-pads.asp"-->
	<!--#include virtual="/lib/crypto.asp"-->
	<%
		
		Dim email
		Dim lista
		Dim opcoes
		Dim KEYTOCRYPT
		KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"
		
		email = Session("user")
		Session.Timeout = TIMEOUT
		secs = CInt((TIMEOUT * 60) + 10)
	
		If StrComp(TypeName(listaPads), "String") = 0 Then
			lista = listaPads
			opcoes = combo
		Else
			lista = join(listaPads)
			opcoes = join(combo)
		End If
	
	
		Dim mensagem
		Dim aux
	
		mensagem = Request.Cookies("mensagem")
		If Len(mensagem) > 0 Then
			aux = mensagem
			mensagem = "<div class=""alert-area""><div class=""alert alert-error"">" & mensagem & "</div> </div>"
		End If
	
	%>
	
	<!-- Basic Page Needs
  ================================================== -->
	<meta charset="utf-8">
	<title>Notejam: New note</title>
	<meta name="description" content="Tela onde criamos notas">
	<meta name="author" content="Diego S. Silva">
	<link rel="icon" href="/diego.silva/assets/img/radioactive.ico">
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
      <h1 class="bold-header"><a href="index.asp" class="header">note<span class="jam">jam:</span></a> <span>New note</span></h1>
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
    <%=mensagem%>
      <form class="note" action="/diego.silva/notejam-vbs-asp-ado/asp/controller/notes/add-a-note.asp" method="POST">
        <label for="name" required>Name</label>
        <input type="text" id="name" name="name" maxlength="100" required autofocus>
        <label for="text">Note</label>
        <textarea id="text" name="text" required></textarea>
        <label for="list">Select Pad</label>
        <select id="list" name="list">
          <option value="0" selected>--------</option>
          <%=opcoes%>
        </select>
        <input type="submit" value="Save">
      </form>
    </div>
    <hr class="footer" />
    <div class="footer">
      <div>Notejam: <strong>VBScript + ASP + ADO</strong> application</div>
      <div>Version by <a href="https://github.com/diegossilva-1995-01-25">Diego S. Silva</a></div>
      <div><a href="https://github.com/komarserjio/notejam">Github</a>, <a href="https://twitter.com/komarserjio">Twitter</a>, based in the app created by <a href="https://github.com/komarserjio/">Serhii Komar</a></div>
    </div>
  </div><!-- container -->
</body>
</html>
<%Response.Cookies("mensagem") = ""%>

