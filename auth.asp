
<%

'response.redirect("ops.asp")

session.lcid = 1046
session("di") = "17/10/2026"
session("df") = "22/10/2026"
session("hi") = 14
session("hf") = "23:59"

'''''''''''''''''''''''''
function LiberaVotacao()
'''''''''''''''''''''''''

if date() < cdate(session("di")) or date() > cdate(session("df")) then
  LiberaVotacao = false
else

  if (date() = cdate(session("di")) and hour(now) < int(session("hi"))) or _
     (date() > cdate(session("df"))) then 'and hour(now) >= int(session("hf"))) then
    LiberaVotacao = false
  else
    LiberaVotacao = true
  end if
end if


end function
%>
