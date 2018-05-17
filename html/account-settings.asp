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
	<!--#include virtual="/lib/crypto.asp"-->
	<%

		Response.CacheControl = "no-store"

		If Session("user") = "" Then
			Response.redirect("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		End If
		
		Dim KEYTOCRYPT
		KEYTOCRYPT = "SOLO_UNA_CHIAVE_PER_CODIFICARE_IL_CORRIERI_ELECTRONICO"

		Dim email
		email = Session("user")
	
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
	<title>Notejam: Account Settings</title>
	<meta name="description" content="Altera a senha">
	<meta name="author" content="Diego S. Silva">
	<link rel="icon" href="/diego.silva/assets/img/radioactive.ico">

	<!-- Mobile Specific Metas
  ================================================== -->
	<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">

	<!-- CSS
  ================================================== -->
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/base.min.css">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/skeleton.min.css">
	<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/skeleton/1.2/layout.css">
	<link rel="stylesheet" href="css/style.css">

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
      <h1 class="bold-header"><a href="index.asp" class="header">note<span class="jam">jam:</span></a> <span>Account Settings</span></h1>
    </div>
    <div class="sixteen columns content-area">
    <%=mensagem%>
      <form class="offset-by-six sign-in" action="/diego.silva/notejam-vbs-asp-ado/asp/controller/login/update-password.asp">
        <label for="current-password">Current password</label>
        <input type="password" id="current-password" name="current-password" maxlength="128" required autofocus>
        <label for="new-password">New password</label>
        <input type="password" id="new-password" name="new-password" maxlength="128" required>
        <label for="confirm-new-password">Confirm new password</label>
        <input type="password" id="confirm-new-password" name="confirm-new-password" maxlength="128" required>
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