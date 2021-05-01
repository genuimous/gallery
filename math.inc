<%
function RoundUp(byval x)
	if x - Round(x) > 0 then
		RoundUp = Round(x) + 1
	else
		RoundUp = Round(x)
	end if
end function
%>
