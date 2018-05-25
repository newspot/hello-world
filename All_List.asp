<!--#include file="../../include/admin_first.inc"-->

<!--#include file="../../../include/adovbs.inc"-->
<!--#include file="../../../include/dbstart.asp"-->
<!--#include file="../../../function/F_COOP_INFO.ASP"-->
<!--#include file="../../../function/F_STR_CUT.ASP"-->
<!--#include file="../../../function/F_NOW_STATE.ASP"-->
<!--#include file="../../../function/F_PAGE_MOVE_ADMIN.ASP"-->
<!--#include file="../../../function/F_STAFF_INFO.ASP"-->
<!--#include file="../../../function/F_CODE_INFO.ASP"-->


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>전체리스트</title>
<!--#include file="../../css/A_style.css"--><!--스타일시트-->
<!--#include file="../../css/variable.inc"--><!--스타일시트-->
<!--#include file="../../js/total.js"--><!--자바스크립트-->
<script language='javascript' src='../../js/formOpen.js'></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
	<%
	Dim pageFormat : pageFormat="Information"
	Dim pageFormatS : pageFormatS="Spage1"
	Dim left_div : left_div = "1"
	%>
	<!--전체를 감싸는 테이블-->
	<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
		<tr>
			<td height="90" align="left" valign="top" colspan="2">
				<!--Top메뉴-->
				<!--#include file="../../include/top.inc" -->
			</td>
		</tr>
		<tr>
			<td width="200" align="left" valign="top" class="left_back">
				<!--Left메뉴-->
				<!--#include file="../../include/Information_left.inc" -->
			</td>
			<td align="left" valign="top">
				<!--Center메뉴-->
				<!--#include file="All_List_core.asp" -->
			</td>
		</tr>
	</table>
</body>
</html>