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
		Response.Cookies("email") = ""
		Session.abandon()

		Dim mensagem
		Dim aux
	
		mensagem = Request.Cookies("mensagem")
		If Len(mensagem) > 0 Then
			
			aux = mensagem
			mensagem = "<div class=""alert-area""><div class=""alert alert-error"">" & aux & "</div> </div>"
		End If
		
		
	%>

	<!-- Basic Page Needs
  ================================================== -->
	<meta charset="utf-8">
	<title>Notejam: Sign Up</title>
	<meta name="description" content="Tela de cadastro de novos usuarios">
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
	<link rel="stylesheet" href="/diego.silva/notejam-vbs-asp-ado/html/css/style.css">

	<!--[if lt IE 9]>
		<script src="http://html5shim.googlecode.com/svn/trunk/html5.js"></script>
	<![endif]-->
</head>
<body>
  <div class="container">
    <div class="sixteen columns">
      <div class="sign-in-out-block">
        <a href="signup.asp">Sign up</a>&nbsp;&nbsp;&nbsp;<a href="signin.asp">Sign in</a>
      </div>
    </div>
    <div class="sixteen columns">
      <h1 class="bold-header"><a href="signin.asp" class="header">note<span class="jam">jam:</span></a> <span> Sign Up</span></h1>
    </div>
    <div class="thirteen columns content-area">
    <%=mensagem%>
      <form class="offset-by-six sign-in" id="form-signup" action="/diego.silva/notejam-vbs-asp-ado/asp/controller/login/login.asp">
        <label for="email">Email</label>
        <input type="email" id="email" name="email" maxlength="75" required autofocus>
        <label for="password">Password</label>
        <input type="password" id="password" name="password" minlength="8" maxlength="128" required>
        <label for="confirm-password">Confirm password</label>
        <input type="password" id="confirm-password" name="confirm-password" minlength="8" maxlength="128" required>
        <input type="submit" value="Sign Up"> or <a href='signin.asp'>Sign in</a>
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
