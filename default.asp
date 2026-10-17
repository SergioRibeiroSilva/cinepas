<!--#include file="auth.asp"-->

<!DOCTYPE html>
<html lang="pt-BR">
<head>
    <meta charset="iso-8859-1">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>CinePAS - Festival de Curtas</title>
    <!-- Fontes premium importadas -->
    <!-- O código existe para garantir uma tipografia elegante, afastando-se das fontes padrão do navegador -->
    <link href="https://fonts.googleapis.com/css2?family=Cinzel:wght@400;700&family=Montserrat:wght@300;500;600&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="style.css">

    <%


    if session("di") = "" then
      response.redirect("ops.asp?msg=Parâmetros para data de votação incorretos! A página do CinePAS deve ser acessada diretamente do site do Santa Rita!")
    end if


    %>
    
</head>
<body>

    <!-- Barra de navegação com os links essenciais -->
    <!-- Existe para que o usuário possa navegar entre a tela de votação e a tela de resultados sem perder o contexto do site -->
    <nav class="navbar">
        <div class="logo">
            <!-- Logotipo da escola inserido no menu superior. -->
            <!-- Existe para manter a marca sempre visível durante a navegação, substituindo o antigo ícone de claquete para reforçar a identidade institucional. -->
            <img src="img/logo-122x122.png" alt="Logo Escola" style="width: 36px; height: auto; filter: drop-shadow(0 0 5px rgba(255, 255, 255, 0.4));">
            <span>CinePAS</span>
        </div>
        <ul class="nav-links">
            <li><a href="default.asp" class="active">Votação</a></li>
            <li><a href="resultado.asp">Resultados Parciais</a></li>
        </ul>
    </nav>

    <!-- Detalhe decorativo de rolo de filme (Film Strip) -->
    <!-- Adicionado para criar uma moldura visual que imediatamente grita "cinema!" logo abaixo do menu -->
    <div class="film-strip"></div>

    <!-- Header principal -->
    <header class="hero">
        <div class="hero-content">
            <!-- Logotipo da instituição no banner. -->
            <!-- Existe para fortalecer o branding (identidade visual) do festival, conectando o evento diretamente à marca da Escola Moderna de forma orgânica. -->
            <img src="img/logo-122x122.png" alt="Logo CinePAS" class="hero-logo">
            <h1>CinePAS 2026</h1>
            <p>O maior festival de curtas do Santa Rita. Assista, viva a magia do cinema e escolha o seu favorito.</p>
            <!-- Período de votação e orientação de voto, conforme o layout anterior (valores vindos de auth.asp) -->
            <p>O período de votação será das <%=session("hi")%>h, do dia <%=session("di")%>, às <%=session("hf")%> do dia <%=session("df")%>.</p>
            <p>Para votar, será necessário informar seu e-mail, lembrando que será validado um voto por e-mail.</p>
        </div>
    </header>

    <!-- Outra moldura de rolo de filme para isolar o título principal -->
    <div class="film-strip"></div>

    <%
    xCaminhoDb = server.MapPath("dados/dbVideo.mdb")
    set ConexaoVideo = Server.CreateObject("ADODB.connection")
    strConexao = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & xCaminhoDb '& ";User Id=Admin;Password=;"
    ConexaoVideo.Open strConexao
    %>

    <!-- Container dos vídeos participantes -->
    <!-- Organiza os vídeos em uma grade. Adicionamos a classe 'cinema-screen' aos cartões para estilizar de acordo com a temática -->
    <!-- Os cartões são gerados a partir dos vídeos cadastrados na tabela "video" do banco dados/dbVideo.mdb -->
    <main class="gallery">

    <%
    set rsVideo = ConexaoVideo.execute("select * from video order by vidCodigo")

    do while not rsVideo.eof
        %>
        <!-- Cartão de Vídeo -->
        <article class="movie-card cinema-screen">
            <div class="video-wrapper">
                <iframe src="https://www.youtube.com/embed/<%=rsVideo("vidCodigoYouTube")%>" frameborder="0" allow="accelerometer; autoplay; clipboard-write; encrypted-media; gyroscope; picture-in-picture" allowfullscreen></iframe>
            </div>
            <div class="card-info">
                <h2><%=server.HTMLEncode(rsVideo("vidTitulo"))%></h2>
                <%
                if LiberaVotacao() then
                %>
                <!-- O botão agora simula visualmente um ingresso dourado de cinema (Golden Ticket) -->
                <button type="button" class="vote-btn ticket" onclick="votar(<%=rsVideo("vidCodigo")%>)">
                    <span class="ticket-text">&#9733; INGRESSO ÚNICO &#9733;<br>Votar Neste Filme</span>
                </button>
                <%
                else
                %>
                <p style="color:#ccc; font-size:0.95rem;">Votação entre <%=session("hi")%>h de <%=session("di")%> e <%=session("hf")%>h de <%=session("df")%></p>
                <%
                end if
                %>
            </div>
        </article>
        <%
        rsVideo.movenext

    loop

    rsVideo.close
    set rsVideo = nothing
    %>

    </main>

    <!-- Detalhe decorativo de rolo de filme no final da página -->
    <div class="film-strip"></div>

    <%
    ConexaoVideo.close
    set ConexaoVideo = nothing
    %>

    <footer style="text-align:center; padding:1.5rem 1rem; color:#aaa; font-size:0.85rem;">
        <a href="http://www.cesrc.com.br" style="color:#ffffff; text-decoration:none;">WWW.CESRC.COM.BR</a><br>
        &copy; Copyright CESRC - Direitos Reservados
    </footer>

    <script>
        function votar(id) {
            window.location.href = 'EfetivaVoto.asp?v=' + id;
        }
    </script>
</body>
</html>
