<%
set ConexaoVideo = Server.CreateObject("ADODB.connection")

'On Error Resume Next
ConexaoVideo.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.MapPath(".") & "\..\dados\dbVideo.mdb"
If Err Then
  Response.Write "erro"
  response.end
End If

set rs = ConexaoVideo.execute("select * from video inner join curtida on video.vidCodigo = curtida.vidCodigo where curVotoConfirmado = 1 order by video.vidCodigo")

do while not rs.eof
  video = rs("video.vidCodigo")
  titulo = rs("vidTitulo")
  votos = 0

  do while video = rs("video.vidCodigo")
    votos = votos + 1
    rs.movenext
    if rs.eof then
      exit do
    end if
  loop

  sResultado = sResultado & titulo & "|" & video & "|" & votos & "_"

loop
rs.close
set rs = nothing

On Error Goto 0

ConexaoVideo.close
set ConexaoVideo = nothing

'retornando uma tripo com: codigo do video | votos separador ;
response.write sResultado

%>
