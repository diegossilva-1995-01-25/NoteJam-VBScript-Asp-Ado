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

		If Session("user") = "" Then
			Response.redirect("https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp")
		End If

		Dim email
		email = Session("user")
	%>
	
	<!-- Basic Page Needs
  ================================================== -->
	<meta charset="utf-8">
	<title>Notejam</title>
	<meta name="description" content="Tela exibida para usuários recém cadastrados.">
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
	<meta http-equiv="refresh" content="300;url=https://blogs.lojcomm.com.br/diego.silva/notejam-vbs-asp-ado/html/signin.asp"/>

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
      <h1 class="bold-header"><a href="index.asp" class="header">note<span class="jam">jam: </span></a> <span>0 notes</span></h1>
    </div>
    <div class="three columns">
      <h4 id="logo">My pads</h4>
      <nav>
        <p class="empty">No pads</p>
      <hr />
      <a href="create-pad.asp">New pad</a>
      </nav>
    </div>
    <div class="thirteen columns content-area">
      <p class="empty">Create your first note.</p>
      <a href="create.asp" class="button">New note</a>
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

