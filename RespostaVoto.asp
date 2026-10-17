
<!--#include file="auth.asp"-->



<%
if LiberaVotacao() = false then
  response.redirect("ops.asp?msg=Parece que não está no período de votação.")
end if
%>

<!DOCTYPE html>
<html >
<head>
  <meta charset="iso-8859-1">

  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
  <link rel="shortcut icon" href="img/logo-122x122.png" type="image/x-icon">
  <meta name="description" content="Confirma&ccedil;&atilde;o de voto - CINEPAS">
  <title>CINEPAS - Confirma&ccedil;&atilde;o de Voto</title>
  <!-- Fontes premium e estilos do layout novo -->
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;700&family=Montserrat:wght@300;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="style.css">
  <style>
    /* Estruturas desta pagina (substituem as classes Bootstrap/Mobirise antigas) */
    .ev-section{padding:2rem 1rem}
    .ev-container{max-width:1000px;margin:0 auto}
    .ev-text{text-align:center;color:#fff;max-width:900px;margin:0 auto}
    .ev-text h2{font-family:'Cinzel',serif;letter-spacing:1px;line-height:1.5}
    .ev-text strong{color:#fff}
    .ev-video{max-width:760px;margin:0 auto 1.5rem}
    .ev-video:hover{transform:none;box-shadow:none}
    .ev-btn{display:inline-block;margin:5px;padding:12px 26px;font-family:'Montserrat',sans-serif;font-size:15px;font-weight:700;text-transform:uppercase;text-decoration:none;background:#315c99;color:#fff;border:2px dashed #fff;border-radius:4px;box-shadow:2px 2px 5px rgba(0,0,0,.5);cursor:pointer;transition:all .3s ease}
    .ev-btn:hover{background:#fff;color:#e42c2d;border-color:#e42c2d;transform:scale(1.03) rotate(-1deg);box-shadow:0 0 15px rgba(255,255,255,.6)}
  </style>

</head>
<body>

  <!-- Barra de navega&ccedil;&atilde;o com os links essenciais (layout novo) -->
  <nav class="navbar">
      <div class="logo">
          <img src="img/logo-122x122.png" alt="Logo Escola" style="width: 36px; height: auto; filter: drop-shadow(0 0 5px rgba(255, 255, 255, 0.4));">
          <span>CinePAS</span>
      </div>
      <ul class="nav-links">
          <li><a href="default.asp" class="active">Vota&ccedil;&atilde;o</a></li>
          <li><a href="resultado.asp">Resultados Parciais</a></li>
      </ul>
  </nav>

  <div class="film-strip"></div>

  <header class="hero" style="min-height: 22vh;">
      <div class="hero-content">
          <img src="img/logo-122x122.png" alt="Logo CinePAS" class="hero-logo">
          <h1>CinePAS 2026</h1>
          <p>Confirma&ccedil;&atilde;o de voto</p>
      </div>
  </header>

  <div class="film-strip"></div>



<!--#include file="md5/md5.asp"-->

<%
xCaminhoDb = server.MapPath("dados/dbVideo.mdb")
set ConexaoVideo = Server.CreateObject("ADODB.connection")
strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDb '& ";User Id=Admin;Password=;"
ConexaoVideo.Open strConexao

titulo = "Confirmação de votação"

if request.QueryString("v") <> "" and request.QueryString("h") <> "" then '''codigo do vÃ­deo e hash do email

	set rsVotoVideo = ConexaoVideo.execute("select * from curtida where curEmail = '" & request.QueryString("h") & "' and curVotoConfirmado = 1" )

	if not rsVotoVideo.eof then
		titulo  = "Ops! O e-mail já realizou uma votação!<br>Lembre-se, um e-mail só pode votar uma vez!"
	else

		set rsVideo = ConexaoVideo.execute("select * from video where vidCodigo = " & request.QueryString("v"))

		if not rsVideo.eof then

			if not isnumeric(request.QueryString("v")) then
				titulo  = "O voto nÃ£o pode ser computado!"
			else

				titulo  = "Parabéns, você acaba de confirmar seu voto para o filme:<br><br>" & server.HTMLEncode(rsVideo("vidTitulo"))

				''''''Neste ponto, gravar o voto ...............
				sq = "UPDATE curtida SET curVotoConfirmado = 1, curData = '" & now & "'"
				sq = sq & " where vidCodigo = " & request.QueryString("v") & " and  curEmail = '" & request.QueryString("h") & "'"

				'response.write sq

				ConexaoVideo.BeginTrans
				ConexaoVideo.execute sq, nAfetado

				if nAfetado = 1 then
					ConexaoVideo.CommitTrans
				else
					titulo = "Ops..........Infelizmente seu voto não foi confirmado!"
					ConexaoVideo.RollBackTrans
				end if
			end if

			''''''''
		end if

		rsVideo.close
		set rsVideo = nothing

	end if

	rsVotoVideo.close
	set rsVotoVideo = nothing

end if
%>

<section class="ev-section" id="content10-4">
    <div class="ev-container">
        <div class="inner-container" style="width: 100%;">
			<br>
            <div class="ev-text">
				&nbsp;
				<h2 style="margin:1px; "><br><%=titulo%></h2>
			</div>
        </div>
    </div>
</section>



<section class="ev-section" id="video1-5">

	<%
	if request.QueryString("v") <> "" and isnumeric(request.QueryString("v")) then

	set rsVideo = ConexaoVideo.execute("select * from video where vidCodigo = " & request.QueryString("v"))

	if not rsVideo.eof then
		%>
		<div class="ev-video movie-card cinema-screen"> 			<div class="video-wrapper"> 				<iframe src="https://www.youtube.com/embed/<%=rsVideo("vidCodigoYouTube")%>" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe> 			</div> 		</div>

		<%
	else
		response.write "<h1 style='text-align:center;'>Vídeo não encontrado!</h1>"
	end if

	rsVideo.close
	set rsVideo = nothing

	else
		response.write "<h1  style='text-align:center;'>Vídeo não informado!</h1>"
	end if



	ConexaoVideo.close
	set ConexaoVideo = nothing
	%>


<h3 style="text-align:center; "><a href="Resultado.asp" class="ev-btn">Acompanhe os resultados</a></h3>


</section>


<!-- Detalhe decorativo de rolo de filme no final da p&aacute;gina -->
<div class="film-strip"></div>

<footer style="text-align:center; padding:1.5rem 1rem; color:#aaa; font-size:0.85rem;">
    <a href="http://www.cesrc.com.br" style="color:#fff; text-decoration:none;">WWW.CESRC.COM.BR</a><br>
    &copy; Copyright CESRC - Direitos Reservados
</footer>


</body>
</html>
