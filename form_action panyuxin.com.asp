{panyuxin.com}

<!--#include file="includes/data_conn.asp"-->
<!--#include file="includes/func2.asp"-->
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
<%
if my_request(Request.QueryString("action"))="save" then
title=Request.Form("title")
content=request.Form("content")
linkman=request.Form("linkman")
tel=request.Form("tel")
address=request.Form("address")
mobiletel=request.Form("mobiletel")
fax=request.Form("fax")
email=request.Form("email")
set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from huang_liuyan",conns,1,3	
			rs.addnew			
			rs("title")	=title
			rs("info")=content
			rs("name")	=linkman			
			rs("tel")=tel
			rs("address")=address
			rs("mobiletel")	=mobiletel			
			rs("fax")=fax
			rs("email")=email					
			rs("addtime")=now			
			rs.update		
			closedata()				
		response.write errbox("Your information has been submitted.","",1)
		response.End()
			end if
			
			
			if my_request(Request.QueryString("action"))="contact" then
title=Request.Form("title")
content=request.Form("content")
linkman=request.Form("linkman")
tel=request.Form("tel")
address=request.Form("address")
mobiletel=request.Form("mobiletel")
fax=request.Form("fax")
email=request.Form("email")
set rs=server.CreateObject("adodb.recordset")
	rs.open "select * from huang_jiameng",conns,1,3	
			rs.addnew			
			rs("title")	=title
			rs("info")=content
			rs("name")	=linkman			
			rs("tel")=tel
			rs("address")=address
			rs("mobiletel")	=mobiletel			
			rs("fax")=fax
			rs("email")=email					
			rs("addtime")=now			
			rs.update		
			closedata()				
		response.write errbox("Your information has been submitted.","",1)
		response.End()
			end if
%>