<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<!--#include file="auth.asp"-->

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="iso-8859-1">
    <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
    <link rel="shortcut icon" href="img/logo-122x122.png" type="image/x-icon">
    <title>Aviso - CinePAS</title>
    
    <link href="https://fonts.googleapis.com/css?family=Montserrat:400,700|Cinzel:400,700" rel="stylesheet">
    <link rel="stylesheet" href="style.css">
    
    <!--#include file="f.asp"-->
    <% Session.LCID = 1046 %>
</head>
<body>

<nav class="ev-navbar">
    <div class="ev-container" style="display:flex; justify-content:space-between; align-items:center;">
        <div class="nav-logo">
            <img src="img/logo-122x122.png" alt="Logo Escola" style="width: 36px; height: auto; filter: drop-shadow(0 0 5px rgba(255, 255, 255, 0.4));">
            <span style="font-family: 'Cinzel', serif; font-weight: bold; font-size: 20px; color: #ffffff; margin-left: 10px;">CinePAS</span>
        </div>
        <div>
            <a href="default.asp" class="ev-btn" style="padding:8px 15px; font-size:12px;">Voltar ao Início</a>
        </div>
    </div>
</nav>

<section class="ev-section">
    <div class="ev-container" style="max-width: 800px; width: 100%; text-align: center; margin-top: 50px;">
        <h2 class="ev-title" style="margin-bottom: 2rem; color: #ffffff;">Ops..!</h2>

    <%if request.querystring("msg") <> "" then %>
      <h1 style="font-size: 24px; line-height: 1.5; color: #fff;"><%=request.querystring("msg")%></h1>      
    <%else%>
      <h1 style="font-size: 24px; line-height: 1.5; color: #fff;">Corrigindo os t&iacute;tulos; Aguarde um minutinho e tente novamente!<br><br> ;)</h1>
    <%end if %>

    </div>
</section>

<!-- Detalhe decorativo de rolo de filme no final da página -->
<div class="film-strip"></div>

<footer style="text-align:center; padding:1.5rem 1rem; color:#aaa; font-size:0.85rem;">
    <a href="http://www.cesrc.com.br" style="color:#ffffff; text-decoration:none;">WWW.CESRC.COM.BR</a><br>
    &copy; Copyright CESRC - Direitos Reservados
</footer>

</body>
</html>
<% response.end %>
