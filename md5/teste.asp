<!--#include file="md5.asp"-->
<%

'''''''''
randomize
sergio = md5("sergio") 
''''''''''''''''''''''''''''''''''''''''

'sergio = md5("sergio") 

response.write sergio & "<br>" 

response.write md5("sergio")

ofeChave = md5("c35a421d2892e241325abd1cb3a307e0")

x = md5("3bffa4ebdf4874e506c2b12405796aa5")




response.write "x=" & x

%>