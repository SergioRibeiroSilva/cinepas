<%
function VerificaEmailDescartavel(ByVal e) 
dim posicaoArrouba, dominio
	
if isnull(e) or trim(e & "") = "" then 
	VerificaEmailDescartavel = false
	exit function 
end if 

posicaoArrouba = instr(e,"@") 
if posicaoArrouba = 0 then
	VerificaEmailDescartavel = false
	exit function
end if

dominio = mid(e,(posicaoArrouba+1)) 

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objFile = objFS.OpenTextFile( Server.MapPath("dacf5d976283753eb4e4e615811c7d7f.txt"))

Do Until objFile.AtEndOfStream

    myLine = objFile.ReadLine	

	if not isnull(myLine) and myLine <> "" then 
	
		if instr(dominio, myLine) > 0 then 				
			
			VerificaEmailDescartavel = true
			
			objFile.Close		
			set objFile = nothing
			set objFS = nothing 
			exit function 
	  
		else 
			VerificaEmailDescartavel = false 
		end if 

	end if 
	
Loop

objFile.Close
set objFile = nothing
set objFS = nothing

end function 
%>
