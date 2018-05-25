<!--Center메뉴-->

<%	' 변수 선언
Dim theUrl, theList : theList="Work/Information/All_List.asp"		' 페이지명

Dim Query
Dim Rs, Rs1
Dim pageCount, totalCount, recordCount, pageSize		' 총 페이지 수, 총 레코드 수, 레코드 카운트, page 사이즈
Dim srchType1, srchType2, srchType3, srchType31, srchType32, srchType33, TempsrchType3, srchStr, srchType_value
Dim GotoPage

pageSize=20
GotoPage = Request("GotoPage")
If GotoPage = "" Then
	GotoPage=1				'처음 페이지가 로딩됬을 경우
End If
srchType1 = Request("srchType1")
srchType2 = Request("srchType2")
srchType3 = Request("srchType3")
srchType31 = Request("srchType31")
srchType32 = Request("srchType32")
srchType33 = Request("srchType33")
srchType_value = Request("srchType_value")

If srchType3<>"" Then
	TempsrchType3=Split(srchType3,"@")
	srchType31=TempsrchType3(0)
	srchType32=TempsrchType3(1)
	srchType33=TempsrchType3(2)
End If

srchType3=srchType31&"@"&srchType32&"@"&srchType33

srchStr = Request("srchStr")
If srchType1="" Then
	srchStr = ""
End If


Dim srch_KEEP_ADDR : srch_KEEP_ADDR = Request("srch_KEEP_ADDR")
Dim srch_CAR_YEAR : srch_CAR_YEAR = Request("srch_CAR_YEAR")
%>

<!--타이틀-->
<table width="1100px" border="0" cellspacing="0" cellpadding="0" bordercolor="<%=bcolor%>">
	<tr><td height="25"></td></tr>
	<tr>
		<td class="ttt">:: 자료 < 전체리스트</td>
	</tr>
	<tr><td><hr size="1" class="hr_01"></td></tr>
</table>

<!--검색-->
<!--#include file="../../core/searchTypeAll.inc"--><!-- Dim srchType1, srchType2, srchType3, srchStr -->
<%
	Query="Select top 500 A.RECEIPT_PATH, A.RECEIPT_TYPE, A.RECEIPT_NO, A.DMGE_NAME, isnull(A.CAR_YEAR,0) as CAR_YEAR, A.DMGE_NUMBER, A.RCPT_DIV, A.RCPT_ID, A.RCPT_PART, A.RCPT_TEAM, A.RQST_NAME, A.RQST_CELL, A.CHARGE, CONVERT(char(10),A.RECEIPT_DATE,120) as RECEIPT_DATE, convert(datetime,A.RECEIPT_DATE) AS RECEIPT_DATE_FULL, A.FLOW_STATE, A.ALLOT_DATE, A.ACDT_NUMBER, A.KEEP_ADDR "
	If srchType_value<>"" Then	' 재생,폐차 설정 시
		Query=Query&" , VALUE_DIV "
		Query=Query&" From VW_DMGE_PEND AS A, T_VALUE_DMGE Where  MEMBER_COMPANY='"&MEMBER_COMPANY&"' and VALUE_DIV LIKE '%"&srchType_value&"%' And a.RECEIPT_NO=T_VALUE_DMGE.RECEIPT_NO" 
	Else
		Query=Query&" From VW_DMGE_PEND AS A Where  MEMBER_COMPANY='"&MEMBER_COMPANY&"'" 
	End If
	If srchType2<>"" Then	' 구분 설정 시
		Query=Query&" and RECEIPT_TYPE='"&srchType2&"'"
	End If
	If srchType31<>"" Then		' 업무단계 설정 시
		Query=Query&" and FLOW_STATE='"&srchType31&"'"
	End If
	If srchType32<>"" Then		' 경매사 설정 시
		Query=Query&" and CHARGE='"&srchType32&"'"
	End If
	If srchType33="COMMON" Then		' 일반건 검색
		Query=Query&" and RCPT_DIV<>'I' "
	ElseIf srchType33<>"" Then		' 제휴사 설정 시
		Query=Query&" and RCPT_ID='"&srchType33&"'"
	End If
	If srchType1<>"" And srchStr<>"" Then ' 검색조건 검색어 설정 시
		Select Case srchType1
			Case "carNumber"			' 차량번호(모델)
				Query=Query&" and DMGE_NUMBER LIKE '%"&srchStr&"%'"
			Case "saleNumber"	' 경매번호
				Query=Query&" and Convert(Int,Right(SALE_NO,5))='"&srchStr&"'"
			Case "carName"				' 차명(제품명)
				Query=Query&" and DMGE_NAME LIKE '%"&srchStr&"%'"
			Case  "OWNER"	' 소유자
				Query=Query&" and OWNER LIKE '%"&srchStr&"%'"
			Case  "OWNER_TELL"	' 소유자
				Query=Query&" and OWNER_TELL LIKE '%"&srchStr&"%'"
			Case "accidentNumber"	' 소유자 연락처
				Query=Query&" and ACDT_NUMBER Like '%"&srchStr&"%'"
			Case  "receiptName"		' 접수자
				Query=Query&" and RQST_NAME LIKE '%"&srchStr&"%'"
			Case  "rqstCell"				' 접수자이동전화
				Query=Query&" and RQST_CELL LIKE '%"&srchStr&"%'"
			Case "receiptNumber"	' 접수번호
				Query=Query&" and Convert(Int,Right(RECEIPT_NO,5))='"&srchStr&"'"
			Case "accidentNumber"	' 보험사고번호
				Query=Query&" and ACDT_NUMBER Like '%"&srchStr&"%'"
		End Select
	End If
	If srch_KEEP_ADDR<>"" Then		' 보관지역 설정 시
		Query=Query&" and KEEP_ADDR='"&srch_KEEP_ADDR&"'"
	End If
	If srch_CAR_YEAR<>"" Then		' 경매사 설정 시
		Query=Query&" and CAR_YEAR like '%"&srch_CAR_YEAR&"%'"
	End If

	Query=Query&" Order By RECEIPT_DATE_FULL Desc"
	Set Rs = Server.CreateObject("ADODB.RecordSet")
	Rs.Open Query, Db_Conn, adOpenStatic
	Rs.PageSize = pageSize
	If Rs.Bof Then
		totalCount = 0
		pageCount = 0
	Else
		totalCount = Rs.RecordCount								' 총 레코드 갯수
		pageCount = Rs.PageCount									' 총 페이지 갯수
		Rs.absolutepage = GotoPage
	End If
%>
<table width="1000px" border="0" cellspacing="0" cellpadding="0" bordercolor="<%=bcolor%>">
	<tr>
		<td align="left"><strong>Total</strong> : <%=totalCount%></td>
		<td align="right"><%=GotoPage%> / <%=pageCount%> pages</td>
	</tr>
</table>

<table class="List_A">
	<tr>
		<th width="40px">경매사</th>
		<th width="86px">구분</th>
		<th width="90px">접수번호</th>
		<th width="100px">차명<br>(제품명)</th>
		<th width="80px">년식</th>
		<th width="60px">차량번호<br>(모델)</th>
		<th width="44px">소속</th>
		<th width="65px">접수자</th>
		<th width="70px">이동<br>전화</th>
		<th width="70px">내부문서</th>
		<th width="40px">외부<br>문서</th>
		<th width="50px">업무<br>단계</th>
		<th width="70px">입찰<br>종료일</th>
		<th width="25px">처리<br>일수</th>
		<th width="25px">보관<br>지역</th>
	</tr>
<%
If Not Rs.BOF Then
	recordCount=0
	Do Until Rs.EOF Or CInt(recordCount) >= Cint(pageSize)
	recordCount = recordCount + 1		' 현재 찍는 게시물이 몇번째인지

	Dim VALUE_DIV, ISSUE_SET
	Query="Select VALUE_DIV, ISSUE_SET From T_VALUE_DMGE Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"'"
	Set Rs1=Server.CreateObject("ADODB.RecordSet")
	Rs1.Open Query, db_Conn, adOpenStatic
	If Rs1.BOF Then
		VALUE_DIV = ""
		ISSUE_SET = ""
	Else	' 데이터가 존재하는 경우
		VALUE_DIV=Rs1("VALUE_DIV")
		ISSUE_SET=Rs1("ISSUE_SET")
	End If 
	Rs1.close
	Set Rs1=Nothing

	Dim IsDb2, CALC_DATE, APRV_OK, UNCO_MONEY
	Query="Select UNIT_NO, CALC_DATE, APRV_OK, (UNCO_FEES + UNCO_VAT) As UNCO_MONEY From T_PROC_CALC Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"' And APRV_OK='Y'"
	Set Rs1=Server.CreateObject("ADODB.RecordSet")
	Rs1.Open Query, db_Conn, adOpenStatic
	If Rs1.BOF Then
		IsDb2 = ""
		CALC_DATE = ""
		APRV_OK = ""
		UNCO_MONEY = 0
	Else	' 데이터가 존재하는 경우
		IsDb2 = "Y"
		CALC_DATE = Rs1("CALC_DATE")
		APRV_OK = Rs1("APRV_OK")
		UNCO_MONEY = Rs1("UNCO_MONEY")
	End If 
	Rs1.close
	Set Rs1=Nothing

	If IsNull(UNCO_MONEY) Or UNCO_MONEY="" Then 
		UNCO_MONEY = 0
	End If 

	Dim END_DATE
	Query="Select RECORD_DATE From T_PROC_END Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"'"
	Set Rs1=Server.CreateObject("ADODB.RecordSet")
	Rs1.Open Query, db_Conn, adOpenStatic
	If Rs1.BOF Then
		END_DATE = "2010-01-01"
	Else	' 데이터가 존재하는 경우
		END_DATE = Rs1("RECORD_DATE")
	End If 
	Rs1.close
	Set Rs1=Nothing

	Dim SALE_NO, SALE_END, SALE_STATE, SPOT_TRNS, SS08_DIRECT
	Query="Select Top 1 SALE_NO, SALE_END, SALE_STATE, SPOT_TRNS, SS08_DIRECT From T_SALE_DMGE Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"' Order By SALE_NO Desc"
	Set Rs1=Server.CreateObject("ADODB.RecordSet")
	Rs1.Open Query, db_Conn, adOpenStatic
	If Rs1.BOF Then
		SALE_END = "-"
		SALE_STATE = ""
		SPOT_TRNS = ""
		SALE_NO=""
		SS08_DIRECT=""
	Else	' 데이터가 존재하는 경우
		SALE_END = Rs1("SALE_END")
		SALE_STATE = Rs1("SALE_STATE")
		SPOT_TRNS = Rs1("SPOT_TRNS")
		SALE_NO=Rs1("SALE_NO")
		SS08_DIRECT=Rs1("SS08_DIRECT")
	End If 
	Rs1.close
	Set Rs1=Nothing
	Dim CAR_YEAR
	If Trim(Rs("CAR_YEAR"))=""  Then 
		CAR_YEAR = 0
	Else
		CAR_YEAR = Rs("CAR_YEAR")		
	End If 
%>
	<tr align="center"<% If Rs("RECEIPT_PATH")="A" Then : Response.Write " bgcolor='#fff799'" : End If %>>
		<td><font style="font-size:11px;"><%Call F_STAFF_INFO(Rs("CHARGE"),"STAFF_NAME","FORMLESS")%></font></td><!-- 경매사 -->

		<td width="60px"><%If InStr(VALUE_DIV, "001")<>0 Then %><img src="../images/001_btn.gif" alt="재생"> <%End If %> <%If InStr(VALUE_DIV, "002")<>0 Then %><img src="../images/002_btn.gif" alt="폐차"> <%End If %><%If InStr(VALUE_DIV, "012")<>0 Then %><img src="../images/012_btn.gif" alt="수출"> <%End If %><%If InStr(VALUE_DIV, "003")<>0 Then %><img src="../images/003_btn.gif" alt="기타"> <%End If %><%If InStr(VALUE_DIV, "004")<>0 Then %><img src="../images/004_btn.gif" alt="부품"> <%End If %>
			<font style="font-size:11px; font-weight:bold; ">
				<%If VALUE_DIV="" Then %>입력 전<%End If %>
			</font>
	
		</td><!-- 구분 -->
		<td><font style="font-size:11px;"><a href="All_View_<%=Rs("RECEIPT_TYPE")%>.asp?RECEIPT_PATH=<%=Rs("RECEIPT_PATH")%>&RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&FLOW_STATE=<%=Rs("FLOW_STATE")%>&srchType1=<%=srchType1%>&srchType2=<%=srchType2%>&srchType3=<%=srchType3%>&srchStr=<%=srchStr%>&theList=<%=theList%>&GotoPage=<%=GotoPage%>" style="cursor:pointer;"><%=Left(Rs("RECEIPT_NO"), 8)&Right(Rs("RECEIPT_NO"), 6)%></a></font></td><!-- 접수번호 -->
		<td><font style="font-size:11px;"><%=cutstr2(Replace(Rs("DMGE_NAME"),"$$",""),23)%></font></td><!-- 차명(제품명) -->
		<td style="font-size:11px;">
			<%=CAR_YEAR%><br>
			<font style="font-size:11px; font-weight:bold; ">
				(<%If CAR_YEAR=0 Then %>입력 전<%ElseIf CDbl(year(now()))-3<CDbl(CAR_YEAR) Then %>3년이내<%ElseIf CDbl(year(now()))-5<CDbl(CAR_YEAR) Then %>3년~5년이내<%Else %>5년 이상<%End If %>)
			</font>
		</td><!-- 년식 -->


		<td><font style="font-size:11px;"><%=cutstr2(Rs("DMGE_NUMBER"),12)%></font></td><!-- 차량번호(모델) -->
		<td><font style="font-size:11px;"><% If Rs("RCPT_DIV")="I" Then %><%Call F_COOP_INFO(Rs("RCPT_ID"),Rs("RCPT_PART"),Rs("RCPT_TEAM"),"COOP_NAME","CUT_STR4")%><% ElseIf Rs("RCPT_DIV")="S" Then %>1종<% ElseIf Rs("RCPT_DIV")="G" Then %> 그·회<% Else %>-<% End If %></font></td><!-- 소속 -->
		<td><font style="font-size:11px;"><%=cutstr2(Rs("RQST_NAME"),8)%></font></td><!-- 접수자 -->
		<td><font style="font-size:11px;"><%=Rs("RQST_CELL")%></font></td><!-- 이동전화 -->



		<td align="left">
		<img alt="접수지" src="../../images/btn_rec.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/form_receipt_<%=Rs("RECEIPT_TYPE")%>.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','damage_open','width=750, height=700, scrollbars=yes, align=center');">


		<img alt="HISTORY REPORT" src="../../images/btn_HR.gif" style="cursor:pointer;" onClick="window.open('../../../Form/History_Report_Admin.asp?SALE_NO=<%=SALE_NO%>&RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','History_Report','width=760, height=800, scrollbars=yes, align=center');">

		<% If ISSUE_SET="Y" Then %> <% If (Rs("RECEIPT_TYPE")="S" Or Rs("RECEIPT_TYPE")="P") Then %><img alt="평가서" src="../../images/btn_Valu.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Valuation_S_Admin.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','Valuation','width=750, height=700, scrollbars=yes, align=center');"><% Else %><img alt="평가서" src="../../images/btn_Valu.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Valuation_Admin.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','Valuation','width=750, height=700, scrollbars=yes, align=center');"><% End If %><% End If %>
		<% If Rs("RCPT_ID")="LT03" Or Rs("RCPT_ID")="SD02" Or Rs("RCPT_ID")="KB40" Or Rs("RCPT_ID")="RENT" Or Rs("RCPT_ID")="TAXI01" Or Rs("RCPT_ID")="BUS" Or Rs("RCPT_ID")="KW42" Or Rs("RCPT_ID")="SY05" Or Rs("RCPT_ID")="GR04" Or Rs("RCPT_ID")="NMCB"  Or Rs("RCPT_ID")="ETC" Or Rs("RCPT_DIV")="C" Or Rs("RCPT_DIV")="G" Then		' 비딩보험사이면 Approve1_Edit_2015.asp로 보낸다 : 2015-05-01부터 시행 %>
			 <img alt="품의서" src="../../images/btn_Consul.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Consultation<% If Rs("RECEIPT_TYPE")="S" Then %>_S<%ElseIf Rs("RECEIPT_TYPE")="P" Then %>_P<% End If %>_Admin_2015.asp?SALE_NO=<%=SALE_NO%>&RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','Consultation','width=750, height=700, scrollbars=yes, align=center');"><%' End If %>
		<% Else %>
			 <img alt="품의서" src="../../images/btn_Consul.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Consultation<% If Rs("RECEIPT_TYPE")="S" Then %>_S<%ElseIf Rs("RECEIPT_TYPE")="P" Then %>_P<% End If %>_Admin.asp?SALE_NO=<%=SALE_NO%>&RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>','Consultation','width=750, height=700, scrollbars=yes, align=center');"><%' End If %>
		<% End If %>

		<% If IsDb2="Y" Then %>
				<img alt="정산서" src="../../images/btn_Calcul.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Calculate_Admin_Benefit.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','Calculate','width=750, height=700, scrollbars=yes, align=center');">
		<% End If %>
		
		<%If Rs("RECEIPT_PATH")="A" Then %>
			<%If APRV_OK="Y" Then %>
				<img alt="거래증명서" src="../../images/AV.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Certificate_KOCAX.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','KOCAX','width=750, height=700, scrollbars=yes, align=center');">
			<%End If %>
		<%End If %>			
		</td><!-- 내부문서 -->

		<td align="left">
			<%If Rs("RECEIPT_PATH")="B" Then %>
				<img alt="매도주문서" src="../../images/btn_appli_C.gif" style="cursor:pointer;" onClick="openApplication('<%=Rs("RECEIPT_NO")%>', '<%=Rs("RECEIPT_TYPE")%>');"><%' If APRV_OK="Y" Then %>
				<% If IsDb2="Y" Then %>
					<% If Rs("RCPT_ID")="SS08" or Rs("RCPT_ID")="HD09"  Then	' 삼성화재비딩건%> <img alt="환입내역서" src="../../images/btn_Calcul_C.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Calculate_Coop_SS08.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','Calculate','width=750, height=700, scrollbars=yes, align=center');">
					<% ElseIf Rs("RCPT_ID")="KODT" Then 	'개인택시 %> <img alt="환입내역서" src="../../images/btn_Calcul_C.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Calculate_Coop_KODT.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','Calculate','width=750, height=700, scrollbars=yes, align=center');">

					<% Else %> <img alt="환입내역서" src="../../images/btn_Calcul_C.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Calculate_Coop_110301.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','Calculate','width=750, height=700, scrollbars=yes, align=center');">
					<% End If %>
				<% End If %>
			<%Else %>
				<img alt="매도주문서" src="../../images/btn_appli_C.gif" style="cursor:pointer;" onClick="openApplication('<%=Rs("RECEIPT_NO")%>', '<%=Rs("RECEIPT_TYPE")%>');">
				<%If APRV_OK="Y" Then %>
					<img alt="거래증명서" src="../../images/AV.gif" style="cursor:pointer;" onClick="javascript:window.open('../../../form/Certificate_KOCAX_Coop.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&RECEIPT_TYPE=<%=Rs("RECEIPT_TYPE")%>&SALE_NO=<%=SALE_NO%>','KOCAX','width=750, height=700, scrollbars=yes, align=center');">
				<%End If %>
			<%End If %>
		</td><!-- 외부문서 -->
				
		<td><font style="font-size:11px;"><%Call F_NOW_STATE(Rs("FLOW_STATE"),"FORMLESS")%></font></td><!-- 업무단계 -->
		<td<% If SS08_DIRECT="Y" Then %>
		bgcolor="#FDEAED" style="cursor:pointer;" onClick="javascript:window.open('../../../Form/Deal_Admin_SUB_SS08.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&SALE_NO=<%=SALE_NO%>','Deal_Admin','width=700, height=700, scrollbars=yes, status=0, align=center');" 
		<% ElseIf SALE_STATE="D" Then %> bgcolor="#FDEAED" style="cursor:pointer;" onClick="javascript:window.open('../../../Form/Deal_Admin.asp?RECEIPT_NO=<%=Rs("RECEIPT_NO")%>&SALE_NO=<%=SALE_NO%>','Deal_Admin','width=700, height=700, scrollbars=yes, status=0, align=center');"
		<% End If %>>
			<font style="font-size:11px;"><%=Mid(SALE_END,3,8)%><% If SALE_STATE="Y" Then %> (<% If SPOT_TRNS="002" Then %><font style="color:red;">경매</font><% ElseIf SPOT_TRNS="003" Then %><font style="color:blue;">공매</font><% ElseIf SPOT_TRNS="004" Then %><font style="color:purple;">당일</font><% End If %>)<% ElseIf SALE_STATE="N" Then %> (취소)<% ElseIf SALE_STATE="D" Then %><span title="즉시처리 사유서를 발급 받으세요!!!"> (즉시)</span><% ElseIf SALE_STATE="U" Then %> (유찰)<% End If %></font>
		<%
		If Rs("FLOW_STATE")="B2" Then
			' 시작가와 최종가 가져오기
			Query="Select START_MONEY, TOP_MONEY From VW_SALE_STATE_Y Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"' And SALE_NO='"&SALE_NO&"'"
'			response.write "q : "&Query&"<br>"
			Set Rs1=Server.CreateObject("ADODB.RecordSet")
			Rs1.Open Query, db_Conn, adOpenStatic

			If Not Rs1.Bof Then
				Response.Write "<br>"&CDbl(Rs1("START_MONEY"))/10000&"만&nbsp;/&nbsp;"&CDbl(Rs1("TOP_MONEY"))/10000&"만"
			End If

			Rs1.close
			Set Rs1=Nothing
		End If
		%>
		</td>
		<td><%
		If Rs("FLOW_STATE")="D0" Then
			Query="Select DISQ_DATE From T_PROC_DISQ Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"'"
			Set Rs1=Server.CreateObject("ADODB.RecordSet")
			Rs1.Open Query, db_Conn, adOpenStatic
			If Rs1.BOF Then
				Response.write "취소"
			Else	' 데이터가 존재하는 경우
				Response.write DateDiff("d",DateValue(Rs("RECEIPT_DATE")),Rs1("DISQ_DATE"))+1
			End If 
			Rs1.close
			Set Rs1=Nothing
		ElseIf InStr(Rs("FLOW_STATE"), "E")<>0 Then
			Query="Select END_DATE, END_DATE2 From T_PROC_END Where RECEIPT_NO='"&Rs("RECEIPT_NO")&"'"
			Set Rs1=Server.CreateObject("ADODB.RecordSet")
			Rs1.Open Query, db_Conn, adOpenStatic
			If Rs1.BOF Then
				Response.write "확인"
			Else	' 데이터가 존재하는 경우
				If IsNull(Rs1("END_DATE")) Then
					Response.write DateDiff("d",DateValue(Rs("RECEIPT_DATE")),Rs1("END_DATE2"))+1
				Else
					Response.write DateDiff("d",DateValue(Rs("RECEIPT_DATE")),Rs1("END_DATE"))+1
				End If
			End If 
			Rs1.close
			Set Rs1=Nothing
		Else
			Response.write DateDiff("d",DateValue(Rs("RECEIPT_DATE")),Date())+1
		End If
		%></td><!-- 처리일수 -->
		<td><% Call F_CODE_INFO (Rs("KEEP_ADDR"), "KEEPING_AREA", "FORMLESS") %> </td><!-- 보관지역 -->



	</tr>
<%
	Rs.moveNext
	Loop
	Else
%>
	<tr>
		<td height="25" align="center" colspan="13">등록된 자료가 없습니다.</td>
	</tr>
<%
End If
%>
</table>
<br>
<!--#include file="../../Core/Page_Move.inc" -->