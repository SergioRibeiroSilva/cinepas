<!--#include file="auth.asp"-->
<!--#include file="f.asp"-->
<!--#include file="md5/md5.asp"-->

<%

Response.Buffer = True

'==== CONFIGURACAO SUPABASE ====
Const SUPABASE_URL = "https://nqcthvefowzpgmjumxve.supabase.co"
Const SUPABASE_ANON_KEY = "sb_publishable_HrGWuEx6nSJGlmAVHnAHig_u8slpKcO"

' Valida o token (JWT) do Supabase e retorna o e-mail autenticado.
' Retorna "" se o token for invalido/expirado.
Function VerificaSupabase(token)
	VerificaSupabase = ""
	if trim(token & "") = "" then exit function

	Dim objXmlHttp, reEmail, mEmail
	on error resume next

	Set objXmlHttp = Server.CreateObject("Msxml2.ServerXMLHTTP")
	objXmlHttp.setTimeouts 5000, 5000, 10000, 10000
	objXmlHttp.open "GET", SUPABASE_URL & "/auth/v1/user", False
	objXmlHttp.setRequestHeader "apikey", SUPABASE_ANON_KEY
	objXmlHttp.setRequestHeader "Authorization", "Bearer " & token
	objXmlHttp.send

	if err.number = 0 then
		if objXmlHttp.status = 200 then
			Set reEmail = New RegExp
			reEmail.IgnoreCase = True
			reEmail.Pattern = """email""\s*:\s*""([^""]+)"""
			Set mEmail = reEmail.Execute(objXmlHttp.responseText)
			if mEmail.Count > 0 then
				VerificaSupabase = lcase(trim(mEmail(0).SubMatches(0)))
			end if
			Set mEmail = Nothing
			Set reEmail = Nothing
		end if
	end if

	Set objXmlHttp = Nothing
	err.clear
	on error goto 0
End Function

'==== ROTA DE VERIFICACAO PREVIA ASSINCRONA DO STATUS DE VOTACAO ====
if request("acao") = "verificar" then
	Response.Clear
	Response.ContentType = "application/json"
	Response.CharSet = "windows-1252"
	dim tokenVerif, emailVerif, emailValido, hashVerif, jaVotouVerif, filmeJaVotado
	tokenVerif = request("supabase_token")
	emailVerif = lcase(trim(request("email") & ""))
	jaVotouVerif = false
	filmeJaVotado = ""
	
	emailValido = VerificaSupabase(tokenVerif)
	if emailValido <> "" and (emailVerif = "" or emailVerif = emailValido) then
		hashVerif = md5(emailValido)
		
		dim xCaminhoDbV, ConexaoVideoV, rsCurtidaV
		xCaminhoDbV = server.MapPath("dados/dbVideo.mdb")
		set ConexaoVideoV = Server.CreateObject("ADODB.connection")
		ConexaoVideoV.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDbV
		
		set rsCurtidaV = ConexaoVideoV.execute("select curtida.*, video.vidTitulo from curtida inner join video on curtida.vidCodigo = video.vidCodigo where curEmail = '" & hashVerif & "' and curVotoConfirmado = 1")
		if not rsCurtidaV.eof then
			jaVotouVerif = true
			filmeJaVotado = rsCurtidaV("vidTitulo") & ""
		end if
		rsCurtidaV.close
		set rsCurtidaV = nothing
		ConexaoVideoV.close
		set ConexaoVideoV = nothing
		
		if jaVotouVerif then
			Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "ok" & chr(34) & "," & chr(34) & "jaVotou" & chr(34) & ":true," & chr(34) & "filmeVotado" & chr(34) & ":" & chr(34) & replace(replace(server.HTMLEncode(filmeJaVotado), "\", "\\"), chr(34), "\" & chr(34)) & chr(34) & "," & chr(34) & "email" & chr(34) & ":" & chr(34) & replace(emailValido, chr(34), "") & chr(34) & "}"
		else
			Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "ok" & chr(34) & "," & chr(34) & "jaVotou" & chr(34) & ":false," & chr(34) & "email" & chr(34) & ":" & chr(34) & replace(emailValido, chr(34), "") & chr(34) & "}"
		end if
	else
		Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "invalido" & chr(34) & "," & chr(34) & "jaVotou" & chr(34) & ":false}"
	end if
	Response.End
end if

'==== ROTA DE REGISTRO PREVIO DO VOTO PENDENTE (MAGIC LINK) ====
if request("acao") = "registrar_pendente" then
	Response.Clear
	Response.ContentType = "application/json"
	Response.CharSet = "windows-1252"
	
	if LiberaVotacao() = false then
		Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "fechado" & chr(34) & "}"
		Response.End
	end if
	
	dim regEmail, regIdVideo, regHash
	regEmail = lcase(trim(request("email") & ""))
	regIdVideo = request("v")
	if not isnumeric(regIdVideo) or trim(regIdVideo & "") = "" then
		regIdVideo = "0"
	end if
	
	if regEmail <> "" and clng(regIdVideo) > 0 and instr(regEmail, "@") > 0 and instr(regEmail, ".") > 0 and not VerificaEmailDescartavel(regEmail) then
		regHash = md5(regEmail)
		
		dim xCaminhoDbP, ConexaoVideoP, rsCurtidaP, sqP
		xCaminhoDbP = server.MapPath("dados/dbVideo.mdb")
		set ConexaoVideoP = Server.CreateObject("ADODB.connection")
		ConexaoVideoP.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDbP
		
		' Verifica se ja existe registro para este hash
		set rsCurtidaP = ConexaoVideoP.execute("select * from curtida where curEmail = '" & regHash & "'")
		if rsCurtidaP.eof then
			' Nao existe registro: insere voto com curVotoConfirmado = 0
			sqP = "INSERT INTO curtida (vidCodigo, curEmail, curReal, curVotoConfirmado, curData) values (" & regIdVideo & ",'" & regHash & "','" & replace(regEmail,"'","''") & "',0,'" & now & "')"
			on error resume next
			ConexaoVideoP.execute sqP
			on error goto 0
		else
			' Se ja existir mas ainda NAO estiver confirmado (curVotoConfirmado <> 1), atualiza o filme e a data
			if rsCurtidaP("curVotoConfirmado") & "" <> "1" and rsCurtidaP("curVotoConfirmado") & "" <> "-1" and rsCurtidaP("curVotoConfirmado") & "" <> "True" then
				sqP = "UPDATE curtida SET vidCodigo = " & regIdVideo & ", curReal = '" & replace(regEmail,"'","''") & "', curVotoConfirmado = 0, curData = '" & now & "' WHERE curEmail = '" & regHash & "'"
				on error resume next
				ConexaoVideoP.execute sqP
				on error goto 0
			end if
		end if
		rsCurtidaP.close
		set rsCurtidaP = nothing
		ConexaoVideoP.close
		set ConexaoVideoP = nothing
		
		Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "ok" & chr(34) & "}"
	else
		Response.Write "{" & chr(34) & "status" & chr(34) & ":" & chr(34) & "invalido" & chr(34) & "}"
	end if
	Response.End
end if

if LiberaVotacao() = false then
  response.redirect("ops.asp?msg=Parece que nao estamos no periodo de votacao.")
end if

autorizadoEnvio = "nao"
dim emailSupabase
emailSupabase = ""
if request.form("supabase_token") <> "" then
	emailSupabase = VerificaSupabase(request.form("supabase_token"))
	if emailSupabase <> "" and lcase(trim(request("email"))) = emailSupabase then
		autorizadoEnvio = "sim"
	end if
end if
%>

<!DOCTYPE html>
<html>
<head>
  <meta charset="iso-8859-1">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
  <link rel="shortcut icon" href="img/logo-122x122.png" type="image/x-icon">
  <meta name="description" content="Vota&ccedil;&atilde;o CINEPAS">
  <title>CINEPAS - Vota&ccedil;&atilde;o</title>
  <!-- Fontes premium e estilos do layout novo -->
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;700&family=Montserrat:wght@300;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="style.css">
  <style>
    /* Estruturas desta pagina (substituem as classes Bootstrap/Mobirise antigas) */
    .ev-section{padding:2rem 1rem}
    .ev-container{max-width:1000px;margin:0 auto}
    .ev-center{text-align:center}
    .ev-text{text-align:center;color:#fff;max-width:900px;margin:0 auto}
    .ev-text h1,.ev-text h2,.ev-container h2{font-family:'Cinzel',serif;letter-spacing:1px;line-height:1.4}
    .ev-text strong{color:#fff}
    .ev-row{display:flex;justify-content:center}
    .ev-col{width:100%;max-width:760px}
    .ev-title{text-align:center;margin-top:1rem}
    .ev-subtitle{font-family:'Cinzel',serif;letter-spacing:1px;padding-bottom:.5rem}
    .ev-video{max-width:760px;margin:0 auto 1.5rem}
    .ev-video:hover{transform:none;box-shadow:none}
    /* Botoes no estilo ticket dourado */
    .ev-btn{display:inline-block;margin:5px;padding:12px 26px;font-family:'Montserrat',sans-serif;font-size:15px;font-weight:700;text-transform:uppercase;text-decoration:none;background:#315c99;color:#fff;border:2px dashed #fff;border-radius:4px;box-shadow:2px 2px 5px rgba(0,0,0,.5);cursor:pointer;transition:all .3s ease}
    .ev-btn:hover{background:#fff;color:#e42c2d;border-color:#e42c2d;transform:scale(1.03) rotate(-1deg);box-shadow:0 0 15px rgba(255,255,255,.6)}
    .ev-btn.ev-btn-sec{background:transparent;color:#ccc;border:1px solid #555;box-shadow:none}
    .ev-btn.ev-btn-sec:hover{color:#e42c2d;border-color:#e42c2d;transform:none}
    .ev-btn.ev-btn-sm{padding:9px 18px;font-size:13px}
    .ev-mt{margin-top:1rem}
    .auth-confirm{background:#315c99;color:#fff;border:2px dashed #fff;border-radius:4px;font-weight:700;font-size:17px;text-transform:uppercase;box-shadow:2px 2px 5px rgba(0,0,0,.5)}
    .auth-confirm:hover{background:#fff;color:#e42c2d;border-color:#e42c2d}
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

<%
if request("email") <> "" and VerificaEmailDescartavel(request("email")) then
%>
	<section class="ev-section" id="content10-4">
		<div class="ev-container">
			<div class="inner-container" style="width: 100%;">
				<br>
				<div class="ev-text">
					&nbsp;
					<br>
					<h1 style="margin:1px;color:white;"><br>Por favor, utilize um tipo de e-mail válido.</h1>
					IP: <%=Request.ServerVariables("REMOTE_ADDR")%>
				</div>
			</div>
		</div>
	</section>
<%
	response.end
end if

xCaminhoDb = server.MapPath("dados/dbVideo.mdb")
set ConexaoVideo = Server.CreateObject("ADODB.connection")
strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDb
ConexaoVideo.Open strConexao

dim idVideo, tituloVideo, youtubeIdVideo
idVideo = request("v")
if not isnumeric(idVideo) or trim(idVideo & "") = "" then 
	idVideo = "0"
end if

tituloVideo = ""
youtubeIdVideo = ""

if clng(idVideo) > 0 then
	set rsVideo = ConexaoVideo.execute("select * from video where vidCodigo = " & idVideo)
	if not rsVideo.eof then
		tituloVideo = rsVideo("vidTitulo")
		youtubeIdVideo = rsVideo("vidCodigoYouTube")
	end if
	rsVideo.close
	set rsVideo = nothing
end if
%>

<%
'=======================================================================
' CASO 1: USUARIO AUTENTICADO VIA SUPABASE (PROCESSAMENTO AUTOMATICO DO VOTO)
'=======================================================================
if autorizadoEnvio = "sim" then
	dim hash, jaVotou, votoGravadoComSucesso, tituloFilmeVotadoAnteriormente
	jaVotou = false
	votoGravadoComSucesso = false
	
	if clng(idVideo) = 0 or tituloVideo = "" then
%>
		<section class="ev-section">
			<div class="ev-container ev-center" style="color:#fff;">
				<br><br>
				<h2>Vídeo não encontrado ou parâmetros incorretos!</h2>
				<br>
				<a href="default.asp" class="ev-btn">Voltar para os vídeos</a>
			</div>
		</section>
<%
	else
		randomize
		hash = md5(emailSupabase)
		
		' 1. Verifica se ja existe voto confirmado para este e-mail
		set rsCurtida = ConexaoVideo.execute("select curtida.*, video.vidTitulo from curtida inner join video on curtida.vidCodigo = video.vidCodigo where curEmail = '" & hash & "' and curVotoConfirmado = 1")
		
		if not rsCurtida.eof then
			jaVotou = true
			tituloFilmeVotadoAnteriormente = rsCurtida("vidTitulo")
			rsCurtida.close
			set rsCurtida = nothing
%>
			<!-- Card: Voto Já Registrado -->
			<section class="ev-section">
				<div class="ev-container">
					<div class="inner-container" style="max-width:720px;margin:60px auto;padding:35px 25px;background:#1e3b63;border:1px solid rgba(255,255,255,0.4);box-shadow:0 15px 40px rgba(0,0,0,0.6);text-align:center;color:#fff;">
						<div style="font-size:45px;margin-bottom:10px;">!!</div>
						<h2 style="color:#fff;font-weight:bold;margin-bottom:15px;">Atenção: Este email já possui voto registrado!</h2>
						<p style="font-size:18px;margin-bottom:12px;">O e-mail <strong><%=server.HTMLEncode(emailSupabase)%></strong> já realizou uma votação no CINEPAS.</p>
						<% if tituloFilmeVotadoAnteriormente <> "" then %>
							<div style="background:rgba(255,255,255,0.1);padding:14px;border-radius:12px;margin:15px 0;">
								<span style="font-size:14px;opacity:0.8;">Seu voto registrado foi no filme:</span><br>
								<strong style="font-size:19px;color:#fff;"><%=server.HTMLEncode(tituloFilmeVotadoAnteriormente)%></strong>
							</div>
						<% end if %>
						<p style="font-size:14px;color:#bbb;margin-top:15px;">* Conforme o regulamento, é permitido apenas 1 voto por e-mail.</p>
						
						<div style="margin-top:30px;display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
							<a href="resultado.asp" class="ev-btn">Ver Resultado Parcial</a>
							<a href="default.asp" class="ev-btn ev-btn-sec">Voltar aos Vídeos</a>
							<button type="button" id="btnTrocarConta" class="ev-btn ev-btn-sec" onclick="trocarConta();">Votar com outra conta</button>
						</div>
					</div>
				</div>
			</section>

			<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
			<script type="text/javascript">
				function trocarConta() {
					var sb = window.supabase.createClient("<%=SUPABASE_URL%>", "<%=SUPABASE_ANON_KEY%>");
					sb.auth.signOut().then(function() {
						window.location.href = "EfetivaVoto.asp?v=<%=idVideo%>";
					});
				}
			</script>
<%
		else
			rsCurtida.close
			set rsCurtida = nothing
			
			' 2. Grava diretamente o voto confirmado
			set rsPendente = ConexaoVideo.execute("select * from curtida where curEmail = '" & hash & "'")
			
			if not rsPendente.eof then
				sq = "UPDATE curtida SET vidCodigo = " & idVideo & ", curReal = '" & replace(emailSupabase,"'","''") & "', curVotoConfirmado = 1, curData = '" & now & "' WHERE curEmail = '" & hash & "'"
			else
				sq = "INSERT INTO curtida (vidCodigo, curEmail, curReal, curVotoConfirmado, curData) values (" & idVideo & ",'" & hash & "','" & replace(emailSupabase,"'","''") & "',1,'" & now & "')"
			end if
			rsPendente.close
			set rsPendente = nothing
			
			on error resume next
			ConexaoVideo.BeginTrans
			ConexaoVideo.execute sq, nAfetado
			
			if nAfetado = 1 and err.number = 0 then
				ConexaoVideo.CommitTrans
				votoGravadoComSucesso = true
			else
				ConexaoVideo.RollBackTrans
				votoGravadoComSucesso = false
			end if
			on error goto 0
			
			if votoGravadoComSucesso then
%>
				<!-- Card: Sucesso Imediato -->
				<section class="ev-section">
					<div class="ev-container">
						<div class="inner-container" style="max-width:720px;margin:60px auto;padding:40px 25px;background:#1e3b63;border:1px solid rgba(255,255,255,0.5);box-shadow:0 15px 45px rgba(0,0,0,0.6);text-align:center;color:#fff;">
							<div style="font-size:50px;margin-bottom:10px;">!</div>
							<h2 style="color:#fff;font-weight:bold;margin-bottom:15px;">Voto Confirmado com Sucesso!</h2>
							<p style="font-size:18px;margin-bottom:8px;">Você votou no filme:</p>
							<div style="background:rgba(255,255,255,0.12);padding:16px 20px;margin:15px 0;font-size:22px;font-weight:700;color:#fff;">
								<%=server.HTMLEncode(tituloVideo)%>
							</div>
							<p style="font-size:15px;color:#ccc;margin-top:15px;">
								Identificado como: <strong><%=server.HTMLEncode(emailSupabase)%></strong>
							</p>
							<p style="font-size:14px;color:#aaa;margin-bottom:25px;">Obrigado por participar da votação do CINEPAS!</p>
							
							<div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;">
								<a href="resultado.asp" class="ev-btn">Ver Resultado Parcial</a>
								<a href="default.asp" class="ev-btn ev-btn-sec">Voltar ao Início</a>
							</div>
						</div>
					</div>
				</section>
<%
			else
%>
				<!-- Card: Erro ao gravar -->
				<section class="ev-section">
					<div class="ev-container">
						<div class="inner-container" style="max-width:600px;margin:40px auto;padding:30px;background:#1e3b63;border-radius:15px;text-align:center;color:#fff;">
							<h2 style="color:#ff4d57;">Não foi possível registrar seu voto!</h2>
							<p>Ocorreu uma falha temporária ao gravar as informações. Por favor, tente novamente.</p>
							<a href="EfetivaVoto.asp?v=<%=idVideo%>" class="ev-btn ev-mt">Tentar novamente</a>
						</div>
					</div>
				</section>
<%
			end if
		end if
	end if

'=======================================================================
' CASO 2: EXIBICAO DO VIDEO E TELA DE IDENTIFICACAO (LOGIN SOCIAL / OTP)
'=======================================================================
else
%>

	<% if clng(idVideo) > 0 and tituloVideo <> "" then %>
		<section class="ev-section" id="content10-4">
			<div class="ev-container">
				<div class="inner-container" style="width: 100%;">
					<br>
					<div class="ev-text">
						&nbsp;<br>
						<h2 style="margin:1px;"><br><br>Você está prestes a votar no filme:<br><br>
						<strong><%=server.HTMLEncode(tituloVideo)%></strong></h2>
					</div>
				</div>
			</div>
		</section>

		<section class="ev-section" id="video1-5">
			<div class="ev-video movie-card cinema-screen"> 				<div class="video-wrapper"> 					<iframe src="https://www.youtube.com/embed/<%=youtubeIdVideo%>" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe> 				</div> 			</div>
		</section>

		<section class="ev-section" id="form1-n">
			<div class="ev-container">
				<div class="ev-row">
					<div class="ev-col">
						
						<div class="ev-row">
							<div class="ev-title">
								<h3 class="ev-subtitle" style="font-size:24px;color:#fff;font-weight:700;">
									Para confirmar seu voto, identifique-se abaixo:
								</h3>
								<p style="font-size:15px;color:#ccc;">
									Seu voto ser&aacute; computado com total seguran&ccedil;a.
								</p>
							</div>
						</div>

						<style>
							.auth-card{max-width:520px;margin:15px auto 30px;padding:32px 26px;border-radius:8px;background:#1e3b63;border:4px solid #2e599b;border-bottom:12px solid #e42c2d;box-shadow:0 15px 40px rgba(0,0,0,0.28);text-align:center;color:#ffffff;animation:authFade .4s ease}
							@keyframes authFade{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
							.auth-btn{display:flex;align-items:center;justify-content:center;gap:10px;width:100%;padding:14px 18px;margin:8px 0;border:0;border-radius:4px;font-size:16px;font-weight:600;cursor:pointer;transition:transform .15s ease,box-shadow .2s ease}
							.auth-btn:hover{transform:translateY(-2px);box-shadow:0 8px 20px rgba(0,0,0,.35)}
							.auth-btn:disabled{opacity:.6;cursor:not-allowed;transform:none;box-shadow:none}
							.auth-google{background:#fff;color:#3c4043}
							.auth-magic{background:linear-gradient(135deg,#e42c2d,#9c1718);color:#fff}
							.auth-input{width:100%;padding:13px 15px;margin:6px 0;border-radius:4px;border:1px solid rgba(255,255,255,.2);font-size:16px;background:rgba(255,255,255,.95);color:#222}
							.auth-input:focus{outline:none;border-color:#e42c2d;box-shadow:0 0 0 3px rgba(228,44,45,.35)}
							.auth-sep{display:flex;align-items:center;gap:10px;margin:18px 0;opacity:.85;font-size:14px;color:#94a3b8;}
							.auth-sep:before,.auth-sep:after{content:"";flex:1;height:1px;background:rgba(255,255,255,0.2);}
							.auth-status{min-height:22px;margin-top:14px;font-size:15px;font-weight:500;line-height:1.4;}
							.auth-ok{color:#4ade80;font-weight:600;}.auth-erro{color:#f87171;font-weight:600;}.auth-info{color:#fff;font-weight:600;}
							.auth-loading{display:flex;flex-direction:column;align-items:center;gap:14px;padding:25px;color:#fff;}
							.spinner{width:42px;height:42px;border:4px solid rgba(255,255,255,0.2);border-top-color:#fff;border-radius:50%;animation:spin 0.9s linear infinite;}
							@keyframes spin{to{transform:rotate(360deg)}}
						</style>

						<form id="myForm" class="form-horizontal" method="post" action="EfetivaVoto.asp?v=<%=server.URLEncode(idVideo)%>">
							<input type="hidden" id="v" name="v" value="<%=idVideo%>">
							<input type="hidden" id="email" name="email" value="">
							<input type="hidden" id="supabase_token" name="supabase_token" value="">

							<div class="auth-card" id="authCard">
								
								<!-- Bloco quando o usuario AINDA NAO esta autenticado -->
								<div id="authLogin">
									<button type="button" id="btnGoogle" class="auth-btn auth-google">
										<svg width="20" height="20" viewBox="0 0 48 48" aria-hidden="true"><path fill="#EA4335" d="M24 9.5c3.54 0 6.71 1.22 9.21 3.6l6.85-6.85C35.9 2.38 30.47 0 24 0 14.62 0 6.51 5.38 2.56 13.22l7.98 6.19C12.43 13.72 17.74 9.5 24 9.5z"/><path fill="#4285F4" d="M46.98 24.55c0-1.57-.15-3.09-.38-4.55H24v9.02h12.94c-.58 2.96-2.26 5.48-4.78 7.18l7.73 6c4.51-4.18 7.09-10.36 7.09-17.65z"/><path fill="#FBBC05" d="M10.53 28.59c-.48-1.45-.76-2.99-.76-4.59s.27-3.14.76-4.59l-7.98-6.19C.92 16.46 0 20.12 0 24c0 3.88.92 7.54 2.56 10.78l7.97-6.19z"/><path fill="#34A853" d="M24 48c6.48 0 11.93-2.13 15.89-5.81l-7.73-6c-2.15 1.45-4.92 2.3-8.16 2.3-6.26 0-11.57-4.22-13.47-9.91l-7.98 6.19C6.51 42.62 14.62 48 24 48z"/></svg>
										Entrar com Google e Votar
									</button>
									
									<div class="auth-sep">ou receba um link de acesso no seu e-mail</div>
									
									<input type="email" id="emailLink" class="auth-input" placeholder="seu@email.com" autocomplete="email">
									<button type="button" id="btnMagic" class="auth-btn auth-magic">Enviar link de acesso</button>
								</div>

								<!-- Bloco quando o usuario JA ESTA autenticado e pode votar -->
								<div id="authConectado" style="display:none;padding:8px 4px;">
									<div style="font-size:15px;color:#cbd5e1;margin-bottom:14px;line-height:1.5;">
										Conectado como:<br>
										<strong id="authEmailLogado" style="font-size:18px;color:#fff;word-break:break-all;"></strong>
									</div>
									<button type="button" id="btnConfirmarVoto" class="auth-btn auth-confirm" style="margin-bottom:14px;">
										Confirmar meu voto agora
									</button>
									<div style="display:flex;gap:10px;justify-content:center;margin-top:10px;flex-wrap:wrap;">
										<a href="default.asp" class="ev-btn ev-btn-sm ev-btn-sec">Escolher outro filme</a>
										<button type="button" id="btnLogout" class="ev-btn ev-btn-sm ev-btn-sec">Trocar de conta / Sair</button>
									</div>
								</div>

								<!-- Bloco quando o usuario JA VOTOU anteriormente -->
								<div id="authJaVotou" style="display:none;padding:8px 4px;">
									<div style="font-size:38px;margin-bottom:8px;">!!</div>
									<h4 style="color:#fff;font-weight:bold;margin-bottom:12px;font-size:22px;">Aten&ccedil;&atilde;o: Este email j&aacute; possui voto registrado!</h4>
									<p style="font-size:15px;color:#cbd5e1;margin-bottom:12px;line-height:1.5;">
										O e-mail <strong id="jaVotouEmail" style="color:#fff;word-break:break-all;"></strong> j&aacute; realizou uma vota&ccedil;&atilde;o no CINEPAS.
									</p>
									<div style="background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.3);padding:14px;border-radius:12px;margin:14px 0;">
										<span style="font-size:13px;color:#94a3b8;">Seu voto registrado foi no filme:</span><br>
										<strong id="jaVotouFilme" style="font-size:18px;color:#fff;"></strong>
									</div>
									<p style="font-size:13px;color:#94a3b8;margin-bottom:18px;">* Conforme o regulamento, &eacute; permitido apenas 1 voto por e-mail.</p>
									
									<div style="display:flex;gap:10px;justify-content:center;flex-wrap:wrap;">
										<a href="resultado.asp" class="ev-btn ev-btn-sm">Ver Resultado Parcial</a>
										<a href="default.asp" class="ev-btn ev-btn-sm ev-btn-sec">Voltar aos V&iacute;deos</a>
										<button type="button" id="btnTrocarContaJaVotou" class="ev-btn ev-btn-sm ev-btn-sec">Votar com outra conta</button>
									</div>
								</div>

								<!-- Bloco de processamento automatico (quando autenticado) -->
								<div id="authProcessando" style="display:none;" class="auth-loading">
									<div class="spinner"></div>
									<div id="authProcessandoMsg" style="font-size:16px;font-weight:600;">Gravando seu voto...</div>
								</div>

								<div id="authStatus" class="auth-status"></div>
							</div>

							<input type="hidden" id="tituloFilmeHidden" value="<%=server.HTMLEncode(tituloVideo)%>">

							<script src="https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2"></script>
							<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
							<script type="text/javascript">
							(function () {
								var sb = window.supabase.createClient("<%=SUPABASE_URL%>", "<%=SUPABASE_ANON_KEY%>");
								var urlRetorno = window.location.origin + window.location.pathname + window.location.search;
								var formEnviado = false;
								var promptExibido = false;
								var sessaoAtual = null;

								function el(id) { return document.getElementById(id); }
								function status(msg, tipo) { 
									var s = el("authStatus"); 
									if (s) {
										s.innerHTML = msg || ""; 
										s.className = "auth-status" + (tipo ? " auth-" + tipo : ""); 
									}
								}

								function enviarVoto(session) {
									if (!session || !session.user || !session.user.email || formEnviado) return;
									formEnviado = true;
									el("email").value = session.user.email;
									el("supabase_token").value = session.access_token;
									
									el("authLogin").style.display = "none";
									if (el("authConectado")) el("authConectado").style.display = "none";
									if (el("authJaVotou")) el("authJaVotou").style.display = "none";
									el("authProcessando").style.display = "flex";
									el("authProcessandoMsg").innerHTML = "Identificado como <strong>" + session.user.email + "</strong><br>Computando seu voto...";
									
									setTimeout(function() {
										el("myForm").submit();
									}, 350);
								}

								function exibirJaVotou(filmeVotado, email) {
									el("authLogin").style.display = "none";
									if (el("authConectado")) el("authConectado").style.display = "none";
									if (el("authProcessando")) el("authProcessando").style.display = "none";
									
									if (el("jaVotouEmail")) el("jaVotouEmail").textContent = email;
									if (el("jaVotouFilme")) el("jaVotouFilme").innerHTML = filmeVotado || "Filme já registrado";
									if (el("authJaVotou")) el("authJaVotou").style.display = "block";
									status("", "");
								}

								function confirmarVotoComSweetAlert(session) {
									if (!session || !session.user || !session.user.email || formEnviado) return;
									sessaoAtual = session;
									var tituloFilme = (el("tituloFilmeHidden") ? el("tituloFilmeHidden").value : "");
									
									el("authLogin").style.display = "none";
									if (el("authJaVotou")) el("authJaVotou").style.display = "none";
									if (el("authEmailLogado")) el("authEmailLogado").textContent = session.user.email;
									if (el("authConectado")) el("authConectado").style.display = "block";

									Swal.fire({
										title: 'Confirmar Voto?',
										html: '<div style="font-size:15px;line-height:1.6;color:#1e293b;text-align:center;">' +
											  'Voc&ecirc; est&aacute; conectado como:<br><strong style="color:#0284c7;font-size:16px;">' + session.user.email + '</strong><br><br>' +
											  'Deseja confirmar o seu voto no filme:<br><strong style="font-size:18px;color:#0f172a;">' + tituloFilme + '</strong>?' +
											  '</div>',
										icon: 'question',
										showCancelButton: true,
										confirmButtonColor: '#e42c2d',
										cancelButtonColor: '#64748b',
										confirmButtonText: 'Sim, confirmar voto!',
										cancelButtonText: 'Cancelar',
										allowOutsideClick: false
									}).then(function (result) {
										if (result.isConfirmed) {
											enviarVoto(session);
										} else {
											status("Voto n&atilde;o computado. Voc&ecirc; pode confirmar no bot&atilde;o acima ou trocar de conta.", "info");
										}
									});
								}

								function tratarSessao(session) {
									if (session && session.user && session.user.email && !formEnviado && !promptExibido) {
										promptExibido = true;
										sessaoAtual = session;

										// Consulta assincrona para saber previamente se o usuario ja votou
										fetch("EfetivaVoto.asp?acao=verificar", {
											method: "POST",
											headers: { "Content-Type": "application/x-www-form-urlencoded" },
											body: "supabase_token=" + encodeURIComponent(session.access_token) + "&email=" + encodeURIComponent(session.user.email)
										})
										.then(function(res) { return res.json(); })
										.then(function(dados) {
											if (dados && dados.status === "ok" && dados.jaVotou) {
												// Usuario ja votou -> Mostra direto o aviso de Voto Ja Registrado
												exibirJaVotou(dados.filmeVotado, session.user.email);
											} else {
												// Usuario nao votou -> Abre modal de confirmacao de voto
												confirmarVotoComSweetAlert(session);
											}
										})
										.catch(function(err) {
											console.error(err);
											confirmarVotoComSweetAlert(session);
										});
									}
								}

								// Verifica sessao ativa ou retorno de OAuth / Magic Link
								sb.auth.getSession().then(function (r) { 
									if (r && r.data && r.data.session) {
										tratarSessao(r.data.session);
									}
								});

								sb.auth.onAuthStateChange(function (evento, session) { 
									if (session && (evento === "SIGNED_IN" || evento === "INITIAL_SESSION")) {
										tratarSessao(session);
									}
								});

								el("btnGoogle").addEventListener("click", function () {
									status("Redirecionando para o Google...", "info");
									sb.auth.signInWithOAuth({ 
										provider: "google", 
										options: { redirectTo: urlRetorno } 
									});
								});

								el("btnMagic").addEventListener("click", function () {
									var em = el("emailLink").value.trim();
									if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(em)) { 
										status("Digite um e-mail v&aacute;lido.", "erro"); 
										return; 
									}
									var b = this; 
									b.disabled = true; 
									status("Enviando link de acesso...", "info");

									// Registra voto pendente (curVotoConfirmado = 0) no banco
									fetch("EfetivaVoto.asp?acao=registrar_pendente", {
										method: "POST",
										headers: { "Content-Type": "application/x-www-form-urlencoded" },
										body: "v=" + encodeURIComponent("<%=idVideo%>") + "&email=" + encodeURIComponent(em)
									}).catch(function(err) {
										console.error("Erro ao registrar pendente:", err);
									});

									sb.auth.signInWithOtp({ 
										email: em, 
										options: { emailRedirectTo: urlRetorno } 
									}).then(function (r) {
										b.disabled = false;
										if (r.error) { 
											status("N&atilde;o foi poss&iacute;vel enviar o link: " + r.error.message, "erro"); 
										} else { 
											status("Link enviado! Acesse seu e-mail e clique no link para votar.", "ok"); 
										}
									});
								});

								if (el("btnConfirmarVoto")) {
									el("btnConfirmarVoto").addEventListener("click", function () {
										if (sessaoAtual) {
											confirmarVotoComSweetAlert(sessaoAtual);
										}
									});
								}

								function deslogar() {
									status("Desconectando...", "info");
									sb.auth.signOut().then(function () {
										sessaoAtual = null;
										promptExibido = false;
										if (el("authJaVotou")) el("authJaVotou").style.display = "none";
										if (el("authConectado")) el("authConectado").style.display = "none";
										el("authLogin").style.display = "block";
										status("Voc&ecirc; saiu da conta. Fa&ccedil;a login com outro e-mail para votar.", "info");
									});
								}

								if (el("btnLogout")) {
									el("btnLogout").addEventListener("click", deslogar);
								}

								if (el("btnTrocarContaJaVotou")) {
									el("btnTrocarContaJaVotou").addEventListener("click", deslogar);
								}

							})();
							</script>

							<div class="ev-center">
								<div>
									<p style="font-size:13px;color:#aaa;margin-top:15px;">* O e-mail ser&aacute; utilizado exclusivamente para garantir a unicidade de 1 voto por pessoa.</p>
								</div>
							</div>

						</form>

					</div>
				</div>
			</div>
		</section>
	<% else %>
		<section class="ev-section">
			<div class="ev-container ev-center" style="color:#fff;">
				<br><br>
				<h2>Vídeo não encontrado!</h2>
				<br>
				<a href="default.asp" class="ev-btn">Voltar para os vídeos</a>
			</div>
		</section>
	<% end if %>

<%
end if
%>

<%
ConexaoVideo.close
set ConexaoVideo = nothing
%>

<!-- Detalhe decorativo de rolo de filme no final da p&aacute;gina -->
<div class="film-strip"></div>

<footer style="text-align:center; padding:1.5rem 1rem; color:#aaa; font-size:0.85rem;">
    <a href="http://www.cesrc.com.br" style="color:#fff; text-decoration:none;">WWW.CESRC.COM.BR</a><br>
    &copy; Copyright CESRC - Direitos Reservados
</footer>

</body>
</html>

<%
'-----------------------------------------------------
'Funcoes auxiliares legadas
'-----------------------------------------------------
Function validadeEmail(ByVal strEmail)
	validadeEmail = IsEmail(strEmail)
End Function

Function IsEmail(ByVal strEmail)
	Dim regEx, ResultadoHum, ResultadoDois, ResultadoTres
    Set regEx = New RegExp
    regEx.IgnoreCase = True
    regEx.Global = True

    regEx.Pattern = "[^@\-\.\w]|^[_@\.\-]|[\._\-]{2}|[@\.]{2}|(@)[^@]*"
    ResultadoHum = RegEx.Test(strEmail)
    regEx.Pattern = "@[\w\-]+\."
    ResultadoDois = RegEx.Test(strEmail)
    regEx.Pattern = "\.[a-zA-Z]{2,3}$"
    ResultadoTres = RegEx.Test(strEmail)

    Set regEx = Nothing

    If Not (ResultadoHum) And ResultadoDois And ResultadoTres Then
        IsEmail = True
    Else
        IsEmail = False
    End If

	if InStr(strEmail,"-") > 0 or InStr(strEmail,"www.") > 0 or InStr(strEmail," ") > 0 then
        IsEmail = False
	end if
End Function

Function DateToNull(Arg)
	if not isdate(Arg) or Arg = "" or isnull(Arg) then
		DateToNull = "null"
	else
		DateToNull = "#" & year(Arg) & "/" & month(Arg) & "/" & day(Arg) & "#"
	end if
end function

Function DataExtenso(dData)
	Dim aSemana, aMes
	aSemana = Array("Domingo", "Segunda-feira", "Terça-feira", "Quarta-feira", "Quinta-feira", "Sexta-feira", "Sábado")
	aMes = Array("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro")

	If IsEmpty(dData) or IsNull(dData) Then
		DataExtenso = ""
	Else
		DataExtenso = aSemana(Weekday(dData)-1) & ", " & Day(dData) & " de " & aMes(Month(dData)-1) & " de " & Year(dData)
		If Hour(dData) <> 0 Then
			DataExtenso = DataExtenso & " às " & Hour(dData) & "h" & Minute(dData)
		End If
	End If
End Function
%>
%>
