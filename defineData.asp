<%
session("di") = request.querystring("di")
session("df") = request.querystring("df")
session("hi") = request.querystring("hi")
session("hf") = request.querystring("hf")

response.redirect("http://cesrc.escolamoderna.com.br")
%>
