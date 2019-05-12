<td>
<%
response.write(time)
response.write(date)

l=request.querystring("fname")
response.write("Full name:" & l & "<br/>")

m=request.querystring("sex")
response.write( "Gender:" & m & "<br/>")

n=request.querystring("liked")
response.write("Liked:" & n & "<br/>")

o=request.querystring("comments")
response.write("You like:" & o & "<br/>")

%>
</td>