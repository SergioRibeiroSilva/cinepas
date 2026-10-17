<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>.:Gerador de Senhas:.</TITLE>
<META NAME="Generator" CONTENT="EditPlus">
<META NAME="Author" CONTENT="Wellington Marinheiro">
</HEAD>
<BODY>
<!--#include file="md5.asp"-->
<%
randomize
tam=8 'Define o tamanho da senha
md=md5(now())
aleat=Int(22-tam*Rnd)+1
senha=mid(md, aleat, tam)

mi = md5("sergio") 

'Explicando:
'response.write "md5(now()) = faz a encriptação (hash) da hora e data atual - (32 caracteres)<BR>"
'response.write "Veja: " & md & "<br>"
'response.write "<HR>Int(22-tam*Rnd)+1 = Gera um aleatorio sempre menor que 22<BR>"
'response.write  "Veja: " & aleat & "<br>"
'response.write "<HR>mid(md, aleat, tam) = Do hash gerado pega o tamanho da senha a partir do aleatorio gerado<HR>"

response.write "Confira a senha gerada: " & senha & "<BR>"
response.write "Confira a senha gerada: " & mi & "<BR>"
%>
</BODY>
</HTML>



