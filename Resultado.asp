<!--#include file="auth.asp"-->
<%
Dim votacaoEncerrada
votacaoEncerrada = false

on error resume next
if date() > cdate(session("df")) then
	votacaoEncerrada = true
elseif date() = cdate(session("df")) and isdate(session("df") & " " & session("hf")) then
	if now >= cdate(session("df") & " " & session("hf")) then
		votacaoEncerrada = true
	end if
elseif not LiberaVotacao() and date() >= cdate(session("di")) then
	votacaoEncerrada = true
end if
on error goto 0
%>

<!DOCTYPE html>
<html>
<head>
  <meta charset="iso-8859-1">
  <meta http-equiv="X-UA-Compatible" content="IE=edge">
  <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
  <link rel="shortcut icon" href="img/logo-122x122.png" type="image/x-icon">
  <meta name="description" content="Resultado da Vota&ccedil;&atilde;o - CINEPAS">
  <title>CINEPAS - <% if votacaoEncerrada then %>Resultado Final<% else %>Resultado Parcial<% end if %></title>
  <!-- Fontes premium e estilos do layout novo -->
  <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;700&family=Montserrat:wght@300;500;600;700&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="style.css">

<!--#include file="f.asp"-->

	<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap" rel="stylesheet">
	<style>
		.rp-wrap{max-width:980px;margin:3rem auto 30px;padding:32px 28px;border-radius:0;background:radial-gradient(circle at 10% 0%,#11233e 0%,#1e3b63 55%,#1b3256 100%);color:#e0e0e0;font-family:'Montserrat',sans-serif;box-shadow:0 20px 60px rgba(0,0,0,.6);border:1px solid rgba(228,44,45,.25);text-align:left}
		.rp-sub{font-size:13px;opacity:.85;margin:0 0 22px;text-align:right;color:#ccc}
		.rp-kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:16px;margin-bottom:30px}
		.rp-kpi{padding:18px 20px;border-radius:0;background:rgba(228,44,45,.08);border:1px solid rgba(228,44,45,.2);backdrop-filter:blur(8px);-webkit-backdrop-filter:blur(8px);transition:transform .2s ease,border-color .2s ease}
		.rp-kpi:hover{transform:translateY(-3px);border-color:rgba(228,44,45,.4)}
		.rp-kpi small{display:block;font-size:11px;letter-spacing:.08em;text-transform:uppercase;color:#aaa}
		.rp-kpi strong{display:block;font-size:34px;font-weight:800;margin-top:4px;line-height:1.1}
		.rp-kpi.v strong{color:#e42c2d}.rp-kpi.p strong{color:#315c99}.rp-kpi.t strong{color:#fff}
		.rp-kpi.d{display:flex;align-items:center;gap:16px}
		.rp-donut{flex:0 0 84px;width:84px;height:84px;border-radius:50%;display:flex;align-items:center;justify-content:center;position:relative}
		.rp-donut:after{content:"";position:absolute;inset:11px;border-radius:50%;background:#1e3b63}
		.rp-donut span{position:relative;z-index:1;font-weight:800;font-size:16px;color:#fff}
		.rp-legend{display:flex;flex-wrap:wrap;gap:18px;justify-content:flex-end;font-size:13px;color:#ccc;margin-bottom:12px}
		.rp-dot{display:inline-block;width:10px;height:10px;border-radius:0;margin-right:6px;vertical-align:middle}
		.rp-item{display:grid;grid-template-columns:44px 96px 1fr;gap:16px;align-items:center;padding:14px;border-radius:0;background:rgba(228,44,45,.05);border:1px solid rgba(228,44,45,.14);margin-bottom:12px;transition:background .2s ease,transform .2s ease,border-color .2s ease;animation:rpIn .6s ease both}
		.rp-item:hover{background:rgba(228,44,45,.12);border-color:rgba(228,44,45,.3);transform:translateX(4px)}
		@keyframes rpIn{from{opacity:0}to{opacity:1}}
		.rp-lider{border-color:rgba(255,255,255,.6);box-shadow:0 8px 30px rgba(228,44,45,.25)}
		.rp-pos{width:40px;height:40px;border-radius:0;display:flex;align-items:center;justify-content:center;font-weight:800;background:rgba(228,44,45,.15);color:#ccc}
		.rp-pos.g1{background:linear-gradient(135deg,#315c99,#1b3256);color:#2a2500;box-shadow:0 0 14px rgba(255,255,255,.5)}
		.rp-pos.g2{background:linear-gradient(135deg,#f1f4fa,#9fb3c2);color:#10222b}
		.rp-pos.g3{background:linear-gradient(135deg,#f3c295,#a86738);color:#261202}
		.rp-thumb{width:96px;height:54px;border-radius:0;object-fit:cover;background:#000;display:block}
		.rp-title{display:flex;justify-content:space-between;gap:10px;font-weight:600;font-size:15px;margin-bottom:8px;color:#fff}
		.rp-pct{color:#e42c2d;font-weight:800;white-space:nowrap}
		.rp-track{height:14px;border-radius:0;background:rgba(255,255,255,.08);overflow:hidden;display:flex}
		.rp-bar{height:100%;width:0;transition:width 1.4s cubic-bezier(.2,.8,.2,1)}
		.rp-bar.v{background:linear-gradient(90deg,#9c1718,#e42c2d);box-shadow:0 0 12px rgba(228,44,45,.7)}
		.rp-bar.p{background:repeating-linear-gradient(45deg,#315c99 0 6px,#1e3b63 6px 12px);opacity:.9}
		.rp-nums{font-size:12px;color:#aaa;margin-top:6px}
		.rp-nums strong{color:#fff}
		.rp-total{text-align:center;margin-top:22px;font-size:14px;color:#aaa}
		.rp-total strong{color:#fff}
		@media (max-width:600px){.rp-item{grid-template-columns:36px 1fr}.rp-thumb{display:none}.rp-kpi strong{font-size:28px}}

		/* Botoes de acao no estilo ticket dourado do layout novo */
		.rp-btn{display:inline-flex;align-items:center;gap:10px;margin:6px;padding:14px 34px;font-family:'Montserrat',sans-serif;font-size:15px;font-weight:700;text-transform:uppercase;text-decoration:none;background:#315c99;color:#fff;border:2px dashed #fff;border-radius:4px;box-shadow:2px 2px 5px rgba(0,0,0,.5);cursor:pointer;transition:all .3s ease}
		.rp-btn:hover{background:#fff;color:#e42c2d;border-color:#e42c2d;transform:scale(1.03) rotate(-1deg);box-shadow:0 0 15px rgba(255,255,255,.6)}
		.rp-btn.rp-btn-sec{background:transparent;color:#ccc;border:1px solid #555;font-size:13px;padding:8px 18px;box-shadow:none}
		.rp-btn.rp-btn-sec:hover{background:transparent;color:#e42c2d;border-color:#e42c2d;transform:none;box-shadow:none}

		/* Estilos exclusivos para Impressao / PDF (Fundo Branco, Alta Nitidez) */
		.print-only {
			display: none;
		}

		@media print {
			html, body {
				background: #ffffff !important;
				background-color: #ffffff !important;
				color: #000000 !important;
				font-family: 'Inter', Arial, Helvetica, sans-serif !important;
				margin: 0 !important;
				padding: 0 !important;
			}
			
			/* Oculta tudo que nao for a secao de impressao */
			body > *:not(#secaoImpressao) {
				display: none !important;
			}
			
			.navbar, 
			.mbr-section, 
			section,
			.cid-r59lMuXWja,
			.cid-r59m9CEolT, 
			.cid-r59mbPSEOQ, 
			.rp-wrap, 
			.no-print, 
			#footer7-7, 
			footer, 
			.engine, 
			hr {
				display: none !important;
			}

			#secaoImpressao, .print-only {
				display: block !important;
				visibility: visible !important;
				width: 100% !important;
				max-width: 100% !important;
				margin: 0 !important;
				padding: 10px 15px !important;
				background: #ffffff !important;
				color: #0f172a !important;
			}

			.print-header {
				border-bottom: 2px solid #0f172a;
				padding-bottom: 14px;
				margin-bottom: 20px;
			}

			.print-logo-title {
				display: flex;
				align-items: center;
				gap: 16px;
				margin-bottom: 10px;
			}

			.print-logo {
				width: 60px;
				height: 60px;
				object-fit: contain;
			}

			.print-title {
				font-size: 18px !important;
				font-weight: 800 !important;
				color: #0f172a !important;
				margin: 0 0 3px 0 !important;
				text-transform: uppercase;
				letter-spacing: 0.5px;
				line-height: 1.2;
			}

			.print-subtitle {
				font-size: 15px !important;
				font-weight: 700 !important;
				color: #0284c7 !important;
				margin: 0 !important;
				letter-spacing: 0.5px;
			}

			.print-meta {
				display: flex;
				justify-content: space-between;
				font-size: 12.5px;
				color: #334155;
				margin-top: 10px;
				background: #f8fafc;
				padding: 8px 12px;
				border: 1px solid #e2e8f0;
			}

			.print-table {
				width: 100% !important;
				border-collapse: collapse !important;
				margin-top: 15px !important;
				font-size: 13.5px !important;
			}

			.print-table th {
				background: #0f172a !important;
				color: #ffffff !important;
				padding: 10px 12px !important;
				font-weight: 700 !important;
				text-transform: uppercase;
				font-size: 12px !important;
				letter-spacing: 0.5px;
				border: 1px solid #0f172a !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.print-table td {
				padding: 10px 12px !important;
				border: 1px solid #cbd5e1 !important;
				color: #1e293b !important;
			}

			.print-table tbody tr:nth-child(even) {
				background: #f8fafc !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.print-table tbody tr.rank-1 {
				background: #fefce8 !important;
				font-weight: bold;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.badge-pos {
				display: inline-block;
				padding: 3px 8px;
				border-radius: 4px;
				font-size: 13px;
				font-weight: 800;
			}

			.badge-gold {
				background: #fef08a !important;
				color: #854d0e !important;
				border: 1px solid #facc15 !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.badge-silver {
				background: #e2e8f0 !important;
				color: #334155 !important;
				border: 1px solid #cbd5e1 !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.badge-bronze {
				background: #ffedd5 !important;
				color: #9a3412 !important;
				border: 1px solid #fed7aa !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.badge-normal {
				color: #475569;
			}

			.print-table tfoot td {
				background: #f1f5f9 !important;
				font-weight: 800 !important;
				border-top: 2px solid #0f172a !important;
				-webkit-print-color-adjust: exact !important;
				print-color-adjust: exact !important;
			}

			.print-footer {
				margin-top: 30px;
				padding-top: 14px;
				border-top: 1px solid #cbd5e1;
				text-align: center;
				font-size: 12px;
				color: #64748b;
			}

			.print-footer p { margin: 3px 0; }
			@page {
				size: A4 portrait;
				margin: 12mm 10mm;
			}
		}
	</style>

</head>
<body>

    <!-- Barra de navega&ccedil;&atilde;o com os links essenciais (layout novo) -->
    <nav class="navbar no-print">
        <div class="logo">
            <img src="img/logo-122x122.png" alt="Logo Escola" style="width: 36px; height: auto; filter: drop-shadow(0 0 5px rgba(255, 255, 255, 0.4));">
            <span>CinePAS</span>
        </div>
        <ul class="nav-links">
            <li><a href="default.asp">Vota&ccedil;&atilde;o</a></li>
            <li><a href="resultado.asp" class="active">Resultados Parciais</a></li>
        </ul>
    </nav>

    <div class="film-strip no-print"></div>

    <!-- Titulo da pagina: Resultado Final ou Parcial conforme votacaoEncerrada -->
    <header class="hero no-print" style="min-height: 25vh;">
        <div class="hero-content">
            <img src="img/logo-122x122.png" alt="Logo CinePAS" class="hero-logo">
            <h1><% if votacaoEncerrada then %>Resultado Final<% else %>Resultado Parcial<% end if %></h1>
            <p>CinePAS 2026</p>
        </div>
    </header>

    <div class="film-strip no-print"></div>

<%
xCaminhoDb = server.MapPath("dados/dbVideo.mdb")
set ConexaoVideo = Server.CreateObject("ADODB.connection")
strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDb
ConexaoVideo.Open strConexao

dim contaTudo, nVotoNaoValidado
contaTudo = 0
nVotoNaoValidado = 0

set rsSoma = ConexaoVideo.execute("select * from curtida")
do while not rsSoma.eof
	if rsSoma("curVotoConfirmado") = 1 then
		contaTudo = contaTudo + 1
	end if
	if isnull(rsSoma("curVotoConfirmado")) or rsSoma("curVotoConfirmado") = 0 then
		nVotoNaoValidado = nVotoNaoValidado + 1
	end if
	rsSoma.movenext
loop
rsSoma.close
set rsSoma = nothing
%>

<section class="no-print" id="video1-5">

	<%
	if contaTudo > 0 or nVotoNaoValidado > 0 or votacaoEncerrada then

		totalPendentes = nVotoNaoValidado
		nContaTodos = 0
		nFilmes = 0
		maxFilme = 0
		Dim aTitulo(), aYT(), aValid(), aPend()

		set rsVideo = ConexaoVideo.execute("select * from video order by vidCodigo")

		do while not rsVideo.eof
			nContaUm = 0
			nPend = 0
			set rsSoma = ConexaoVideo.execute("select * from curtida where vidCodigo = " & rsVideo("vidCodigo"))

			do while not rsSoma.eof
				if rsSoma("curVotoConfirmado") = 1 then
					nContaUm = nContaUm + 1
				end if
				if isnull(rsSoma("curVotoConfirmado")) or rsSoma("curVotoConfirmado") = 0 then
					nPend = nPend + 1
				end if
				nContaTodos = nContaTodos + 1
				rsSoma.movenext
			loop
			rsSoma.close
			set rsSoma = nothing

			ReDim Preserve aTitulo(nFilmes)
			ReDim Preserve aYT(nFilmes)
			ReDim Preserve aValid(nFilmes)
			ReDim Preserve aPend(nFilmes)
			aTitulo(nFilmes) = rsVideo("vidTitulo") & ""
			aYT(nFilmes) = rsVideo("vidCodigoYouTube") & ""
			aValid(nFilmes) = nContaUm
			aPend(nFilmes) = nPend
			if (nContaUm + nPend) > maxFilme then maxFilme = nContaUm + nPend
			nFilmes = nFilmes + 1

			rsVideo.movenext
		loop
		rsVideo.close
		set rsVideo = nothing

		if maxFilme = 0 then maxFilme = 1

		' Ordena: mais votos validados primeiro (desempate pelos pendentes)
		for i = 0 to nFilmes - 2
			for j = 0 to nFilmes - 2 - i
				if aValid(j) < aValid(j+1) or (aValid(j) = aValid(j+1) and aPend(j) < aPend(j+1)) then
					tmp = aTitulo(j) : aTitulo(j) = aTitulo(j+1) : aTitulo(j+1) = tmp
					tmp = aYT(j)     : aYT(j)     = aYT(j+1)     : aYT(j+1)     = tmp
					tmp = aValid(j)  : aValid(j)  = aValid(j+1)  : aValid(j+1)  = tmp
					tmp = aPend(j)   : aPend(j)   = aPend(j+1)   : aPend(j+1)   = tmp
				end if
			next
		next

		pctGeral = 0
		if (contaTudo + totalPendentes) > 0 then
			pctGeral = (contaTudo * 100) / (contaTudo + totalPendentes)
		end if
	%>

	<div class="no-print" style="padding: 0 1rem;">
	<div class="rp-wrap">

		<% if votacaoEncerrada then %>
		<p class="rp-sub" style="color:#ffffff;font-weight:600;font-size:14px;">Resultado Final &bull; Vota&ccedil;&atilde;o Encerrada</p>
		<% else %>
		<p class="rp-sub">Pr&eacute;via atualizada em <%=now%></p>
		<% end if %>

		<div class="rp-kpis">
			<div class="rp-kpi v"><small>Votos validados</small><strong class="rp-num" data-n="<%=contaTudo%>"><%=contaTudo%></strong></div>
			<div class="rp-kpi p"><small>Aguardando valida&ccedil;&atilde;o</small><strong class="rp-num" data-n="<%=totalPendentes%>"><%=totalPendentes%></strong></div>
			<div class="rp-kpi t"><small>Filmes concorrendo</small><strong class="rp-num" data-n="<%=nFilmes%>"><%=nFilmes%></strong></div>
			<div class="rp-kpi d">
				<div class="rp-donut" style="background:conic-gradient(#315c99 0 <%=CssNum(pctGeral)%>%,#ffffff <%=CssNum(pctGeral)%>% 100%);"><span><%=FormatNumber(pctGeral,0)%>%</span></div>
				<small>dos votos j&aacute; foram validados</small>
			</div>
		</div>

		<div class="rp-legend">
			<span><i class="rp-dot" style="background:#315c99;"></i>Validados</span>
			<span><i class="rp-dot" style="background:#ffffff;"></i>Aguardando valida&ccedil;&atilde;o</span>
		</div>

		<%
		for i = 0 to nFilmes - 1
			pctFilme = 0
			if contaTudo > 0 then pctFilme = (aValid(i) * 100) / contaTudo
			wV = (aValid(i) * 100) / maxFilme
			wP = (aPend(i) * 100) / maxFilme
			classePos = ""
			if i < 3 then classePos = " g" & (i + 1)
			classeItem = ""
			if i = 0 and aValid(i) > 0 then classeItem = " rp-lider"
		%>
		<div class="rp-item<%=classeItem%>" style="animation-delay:<%=CssNum(i * 0.08)%>s;">
			<div class="rp-pos<%=classePos%>"><%=i + 1%></div>
			<% if aYT(i) <> "" then %>
			<img class="rp-thumb" src="https://img.youtube.com/vi/<%=server.HTMLEncode(aYT(i))%>/mqdefault.jpg" alt="" loading="lazy">
			<% else %>
			<div class="rp-thumb"></div>
			<% end if %>
			<div>
				<div class="rp-title"><span><%=server.HTMLEncode(aTitulo(i))%></span><span class="rp-pct"><%=FormatNumber(pctFilme,1)%>%</span></div>
				<div class="rp-track">
					<div class="rp-bar v" data-w="<%=CssNum(wV)%>" title="<%=aValid(i)%> validado(s)"></div>
					<div class="rp-bar p" data-w="<%=CssNum(wP)%>" title="<%=aPend(i)%> aguardando"></div>
				</div>
				<div class="rp-nums"><strong><%=aValid(i)%></strong> validado(s) &nbsp;&middot;&nbsp; <%=aPend(i)%> aguardando valida&ccedil;&atilde;o</div>
			</div>
		</div>
		<%
		next
		%>

		<div class="rp-total">Votos apurados: <strong><%=nContaTodos%></strong></div>

	</div>

	<!-- Botoes de Acao abaixo do Quadro -->
	<div class="no-print" style="text-align:center; margin: 25px auto 40px;">
		<% if votacaoEncerrada then %>
		<button type="button" onclick="window.print();" class="rp-btn">
			<svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M6 9V2h12v7"></path><path d="M6 18H4a2 2 0 0 1-2-2v-5a2 2 0 0 1 2-2h16a2 2 0 0 1 2 2v5a2 2 0 0 1-2 2h-2"></path><path d="M6 14h12v8H6z"></path></svg>
			Imprimir Resultado Final (PDF)
		</button>
		<% end if %>
		<div style="margin-top: 12px;">
			<a href="default.asp" class="rp-btn rp-btn-sec">Voltar ao In&iacute;cio</a>
		</div>
	</div>

	</div>

	<script type="text/javascript">
	(function () {
		function anima() {
			var b = document.querySelectorAll(".rp-bar");
			for (var i = 0; i < b.length; i++) { b[i].style.width = b[i].getAttribute("data-w") + "%"; }
			var n = document.querySelectorAll(".rp-num");
			for (var k = 0; k < n.length; k++) {
				(function (el) {
					var alvo = parseInt(el.getAttribute("data-n"), 10) || 0, ini = null;
					function passo(t) {
						if (!ini) ini = t;
						var p = Math.min((t - ini) / 1000, 1);
						el.textContent = Math.round(alvo * (1 - Math.pow(1 - p, 3)));
						if (p < 1) requestAnimationFrame(passo);
					}
					requestAnimationFrame(passo);
				})(n[k]);
			}
		}
		if (document.readyState === "complete") { setTimeout(anima, 150); }
		else { window.addEventListener("load", function () { setTimeout(anima, 150); }); }
	})();
	</script>

	<%
	else
		response.write "<h1 style='text-align:center;color:#fff;padding:60px 20px;'>Aguardando in&iacute;cio de vota&ccedil;&atilde;o</h1>"
	end if
	%>

</section>

<% if contaTudo > 0 or votacaoEncerrada then %>
<!-- SECAO EXCLUSIVA PARA IMPRESSAO / PDF (FILHOTE DIRETO DO BODY) -->
<div id="secaoImpressao" class="print-only">
	<div class="print-header">
		<div class="print-logo-title">
			<img src="img/logo-122x122.png" alt="Logo" class="print-logo">
			<div>
				<h1 class="print-title">CENTRO DE ENSINO SANTA RITA DE C&Aacute;SSIA - CINEPAS</h1>
				<h2 class="print-subtitle">RESULTADO FINAL</h2>
			</div>
		</div>
		<div class="print-meta">
			<span><strong>Data/Hora de Apura&ccedil;&atilde;o:</strong> <%=now%></span>
			<span><strong>Total de Votos Validados:</strong> <%=contaTudo%></span>
			<span><strong>Total de Filmes Concorrendo:</strong> <%=nFilmes%></span>
		</div>
	</div>

	<table class="print-table">
		<thead>
			<tr>
				<th style="width: 100px; text-align: center;">Posi&ccedil;&atilde;o</th>
				<th style="text-align: left;">T&iacute;tulo do Filme</th>
				<th style="width: 140px; text-align: center;">Votos Validados</th>
				<th style="width: 130px; text-align: right;">% dos Votos</th>
			</tr>
		</thead>
		<tbody>
			<%
			for i = 0 to nFilmes - 1
				pctFilme = 0
				if contaTudo > 0 then pctFilme = (aValid(i) * 100) / contaTudo
				classeRank = ""
				if i = 0 then classeRank = "rank-1"
				if i = 1 then classeRank = "rank-2"
				if i = 2 then classeRank = "rank-3"
			%>
			<tr class="<%=classeRank%>">
				<td style="text-align: center; font-weight: bold;">
					<% if i = 0 then %>
						<span class="badge-pos badge-gold">1&ordm; Lugar</span>
					<% elseif i = 1 then %>
						<span class="badge-pos badge-silver">2&ordm; Lugar</span>
					<% elseif i = 2 then %>
						<span class="badge-pos badge-bronze">3&ordm; Lugar</span>
					<% else %>
						<span class="badge-pos badge-normal"><%=i + 1%>&ordm; Lugar</span>
					<% end if %>
				</td>
				<td style="font-weight: 600; font-size: 15px;">
					<%=server.HTMLEncode(aTitulo(i))%>
				</td>
				<td style="text-align: center; font-weight: bold; font-size: 15px;">
					<%=aValid(i)%>
				</td>
				<td style="text-align: right; font-weight: bold; font-size: 15px;">
					<%=FormatNumber(pctFilme, 1)%>%
				</td>
			</tr>
			<%
			next
			%>
		</tbody>
		<tfoot>
			<tr>
				<td colspan="2" style="font-weight: bold; text-align: right;">TOTAL DE VOTOS VALIDADOS:</td>
				<td style="text-align: center; font-weight: bold;"><%=contaTudo%></td>
				<td style="text-align: right; font-weight: bold;">100,0%</td>
			</tr>
		</tfoot>
	</table>

	<div class="print-footer">
		<p><strong>Centro de Ensino Santa Rita de C&aacute;ssia &bull; CINEPAS</strong> &bull; WWW.CESRC.COM.BR</p>
		<p>Documento de apura&ccedil;&atilde;o oficial gerado automaticamente pelo Sistema CINEPAS</p>
	</div>
</div>
<% end if %>

<%
ConexaoVideo.close
set ConexaoVideo = nothing
%>

<!-- Detalhe decorativo de rolo de filme no final da p&aacute;gina -->
<div class="film-strip no-print"></div>

<footer class="no-print" style="text-align:center; padding:1.5rem 1rem; color:#aaa; font-size:0.85rem;">
    <a href="http://www.cesrc.com.br" style="color:#315c99; text-decoration:none;">WWW.CESRC.COM.BR</a><br>
    &copy; Copyright CESRC - Direitos Reservados
</footer>

</body>
</html>

<%
' Formata numero com ponto decimal (para uso em CSS, independente do LCID)
Function CssNum(n)
	CssNum = Replace(CStr(Round(n, 2)), ",", ".")
End Function
%>
