<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>

<!--#include file="auth.asp"-->

<!DOCTYPE html>
<html lang="pt-br">
<head>
    <meta charset="iso-8859-1">
    <meta name="viewport" content="width=device-width, initial-scale=1, minimum-scale=1">
    <link rel="shortcut icon" href="img/logo-122x122.png" type="image/x-icon">
    <title>Lista de Votos - CinePAS</title>
    
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
    <div class="ev-container" style="max-width: 1000px; width: 100%; overflow-x: auto;">
        <h2 class="ev-title" style="margin-bottom: 2rem;">Lista de Votos</h2>

<%

xCaminhoDb = server.MapPath("dados/dbVideo.mdb")
set ConexaoVideo = Server.CreateObject("ADODB.connection")
strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDb '& ";User Id=Admin;Password=;"
ConexaoVideo.Open strConexao
%>

	<%

	if  request.QueryString("c") <> "" and isnumeric(request.QueryString("c")) then

		set rs = ConexaoVideo.execute("select * from video inner join curtida on video.vidCodigo = curtida.vidCodigo where video.vidCodigo = " & request.QueryString("c") & " order by vidTitulo, curData")

	else
		set rs = ConexaoVideo.execute("select * from video inner join curtida on video.vidCodigo = curtida.vidCodigo order by vidTitulo, curData")
	end if


	if not rs.eof then

		response.write "<table class='ev-table' style='width:100%; text-align:left; font-size:13px; border-collapse: collapse; margin-top:20px;' align='center'>"

		response.write "<tr style='background-color:rgba(255, 255, 255, 0.1); color:#ffffff; border-bottom:1px solid rgba(255, 255, 255, 0.3);'>"
		response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>#</td>"
		response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Titulo</td>"
		'response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Email</td>"
		response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>E-mail</td>"
		response.write "<td style='text-align:center; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Situa&ccedil;&atilde;o</td>"
		response.write "<td style='text-align:center; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Confirmado<br>no e-mail</td>"
		'response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Data</td>"

		response.write "</tr>"

		vidTitulo = rs("vidTitulo")

		nVotoConfirmado = 0
		nVotoDescartavel = 0
		nVotoValido = 0
		nVotoNaoValidado = 0
		do while not rs.eof

			nConta = nConta + 1

			response.write "<tr>"
			response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & nConta & "</td>"
			response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & server.HTMLEncode(rs("vidTitulo")) & "</td>"

			'response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & rs("curReal") & "</td>"
			response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & mid(rs("curEmail"),1,18) & "</td>"

			'if VerificaEmailDescartavel(rs("curReal"))  then
				'response.write "<td style='text-align:center; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Em anolise</td>"
			'else
				response.write "<td style='text-align:center; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>Ok</td>"
			'end if

			'if VerificaEmailDescartavel(rs("curReal")) then
				'nVotoDescartavel = nVotoDescartavel + 1
				'response.write "<td style='text-align:center;color:red; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>NaO</td>"
			'else
				nVotoValido = nVotoValido + 1

				if rs("curVotoConfirmado") = 0 then
					response.write "<td style='text-align:center;color:red; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>N&Atilde;O</td>"
					nVotoNaoValidado = nVotoNaoValidado + 1
				else
					response.write "<td style='text-align:center;color:green; padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>SIM</td>"
					nVotoConfirmado = nVotoConfirmado + 1
				end if

			'end if



			'response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & month(rs("curData")) & "/" & day(rs("curData")) & "/" & year(rs("curData")) & " " & hour(rs("curData")) & ":" & minute(rs("curData")) & ":" & second(rs("curData")) & "</td>"
	'		response.write "<td style='padding:8px; border-bottom:1px solid rgba(255,255,255,0.1);'>" & rs("curData") & "</td>"
			response.write "</tr>"

			nContaTodos = nContaTodos + 1

			rs.movenext



			if not rs.eof then

				if vidTitulo <> rs("vidTitulo")  then

					response.write "<tr style='background-color:rgba(255,255,255,0.05);color:#fff;'>"
					response.write "<td colspan='8' style='padding:15px 8px; border-bottom:1px solid rgba(255, 255, 255, 0.3); font-weight:bold;'>Total de votos: " & nConta & "<br>V&aacute;lidos: " & nVotoValido
'					response.write "<br>Em anolise: " & nVotoDescartavel
					response.write"<br>Confirmados: " & nVotoConfirmado & "<br>N&atilde;o confirmados: " & nVotoNaoValidado & "</td>"
					response.write "</tr>"

					nVotoConfirmado = 0
					nVotoDescartavel = 0
					nVotoValido = 0
					nConta = 0
					nVotoNaoValidado = 0

					vidTitulo = rs("vidTitulo")
				end if
			end if

		loop

		response.write "<tr style='background-color:rgba(255,255,255,0.05);color:#fff;'>"
		response.write "<td colspan='8' style='padding:15px 8px; border-bottom:1px solid rgba(255, 255, 255, 0.3); font-weight:bold;'>Total de votos: " & nConta & "<br>V&aacute;lidos: " & nVotoValido & "<br>"

'		response.write "Em anolise: " & nVotoDescartavel &"<br>"

		response.write "Confirmados: " & nVotoConfirmado & "<br>N&atilde;o confirmados: " & nVotoNaoValidado & "</td>"

		response.write "</tr>"

		response.write "</table>"

		response.write "<h1 style='text-align:center; margin-top:20px; color:#ffffff;'>" & nContaTodos & "</h1>"

		rs.close
		set rs = nothing

	end if
	%>

<%
ConexaoVideo.close
set ConexaoVideo = nothing
%>

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
