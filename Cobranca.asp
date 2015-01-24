<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
	<head>
		<title>Cobranca</title><LINK href="Boleto.css" type="text/css" rel="stylesheet">
			<%

'Força mudança de layout para formato brasileiro
Session.LCID = 1046

'Cria a conexão com o banco de dados
SET oDB=Server.CreateObject( "ADODB.Connection" )
oDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ProdCobranca.mdb")

'Abre uma tabela de configurações para obter os dados do cedente
SET oConfig=Server.CreateObject( "ADODB.RecordSet" )
oConfig.Open "Configuracoes", oDB, 2, 2

if Request("RUN")<>"" then

	SET oCobranca=Server.CreateObject( "ADODB.RecordSet" )
	oCobranca.Open "Cobrancas", oDB, 2, 2
	
	oCobranca.AddNew()
	oCobranca("Sacado")=Request("Sacado")
	oCobranca("Documento")=Request("Documento")
	oCobranca("Endereco")=Request("Endereco")
	oCobranca("BairroCidade")=Request("BairroCidade")
	oCobranca("CepUF")=Request("CepUF")
	oCobranca("Valor")=cDbl(Request("Valor"))
	oCobranca("Vencimento")=cDate(Request("Vencimento"))
	oCobranca("Demonstrativo")=Request("Demonstrativo")
	oCobranca("Instrucoes")=Request("Instrucoes")
	oCobranca("email")=Request("email")
	oCobranca.Update()

	if request("email")<>"" then
		on error resume next
		'=========================== CDO SYS
		'cria o objeto para o envio de e-mail 
		Set objCDOSYSMail = Server.CreateObject("CDO.Message")
		'cria o objeto para configuração do SMTP 
		Set objCDOSYSCon = Server.CreateObject ("CDO.Configuration")
		'SMTP 
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "localhost"
		'porta do SMTP 
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		'porta do CDO 
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		'timeout 
		objCDOSYSCon.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 30 
		objCDOSYSCon.Fields.update 
		'atualiza a configuração do CDOSYS para o envio do e-mail 
		Set objCDOSYSMail.Configuration = objCDOSYSCon
		'e-mail do remetente 
		objCDOSYSMail.From = "site@boletoasp.com.br"
		'e-mail do destinatário 
		objCDOSYSMail.To = request("email")
		'e-mail do cópia
		'objCDOSYSMail.CC = "fabio@impactro.com.br"
		'assunto da mensagem 
		objCDOSYSMail.Subject = "Cobrança"
		'conteúdo da mensagem 
		objCDOSYSMail.TextBody = "Segue sua cobrança: " & vbCrLf & _
			Request("Demonstrativo") & vbCrLf & _
			"Acesse o link http://www.boletoasp.com.br/GeraBoleto.asp?id=" & oCobranca("NossoNumero") & " para efetuar o pagamento." & vbCrLf & vbCrLf & _
			"Obrigado."
		'para envio da mensagem no formato html altere o TextBody para HtmlBody 
		'objCDOSYSMail.HtmlBody = "Teste do componente CDOSYS"

		'objCDOSYSMail.fields.update
		'envia o e-mail 
		objCDOSYSMail.Send 
		
		if Err<>0 then
			Response.Write(err.Description)
			oConfig.Close()
            oDB.Close()
		end if

		'destrói os objetos 
		Set objCDOSYSMail = Nothing 
		Set objCDOSYSCon = Nothing

	end if
	
	oConfig.Close()
    oDB.Close()
	
	Response.Redirect( "GeraBoleto.asp?id=" & oCobranca("NossoNumero"))

end if

%>
	</head>
	<body>
		<form action="Cobranca.asp" method="post">
			<TABLE id="Table1" cellSpacing="1" cellPadding="1" align="center" border="0">
				<TR>
					<TH colSpan="4">
					Informações do Cedente</TD>
				</TR>
				<TR>
					<TD align="center" colSpan="4" height="10">&nbsp;&nbsp;&nbsp;</TD>
				</TR>
				<TR>
					<TD width="100">Cedente:</TD>
					<TD colSpan="3"><b><%=oConfig("Cedente")%></b></TD>
				</TR>
				<TR>
					<TD>Banco:</TD>
					<TD width="100"><%=oConfig("Banco")%></TD>
					<TD width="100">Cod.Cedente:</TD>
					<TD width="100"><%=oConfig("CodCedente")%></TD>
				</TR>
				<TR>
					<TD>Agência:</TD>
					<TD><%=oConfig("Agencia")%></TD>
					<TD>Modalidade:</TD>
					<TD><%=oConfig("Modalidade")%></TD>
				</TR>
				<TR>
					<TD>Conta:</TD>
					<TD><%=oConfig("Conta")%></TD>
					<TD>Convénio:</TD>
					<TD><%=oConfig("Convenio")%></TD>
				</TR>
				<TR>
					<TD>Carteira:</TD>
					<TD><%=oConfig("Carteira")%></TD>
				</TR>
				<TR>
					<TD colSpan="4" height="10"></TD>
				</TR>
			</TABLE>
			<BR>
			<TABLE id="Table2" cellSpacing="1" cellPadding="1" width="90%" align="center" border="0">
				<TR>
					<TH colSpan="2">
					Informações do Sacado</TD>
				</TR>
				<TR>
					<TD align="center" colSpan="2" height="10"></TD>
				</TR>
				<TR>
					<TD height="26">Sacado:</TD>
					<TD height="26"><INPUT id="Text5" type="text" maxLength="50" size="50" name="Sacado" DESIGNTIMEDRAGDROP="421"></TD>
				</TR>
				<TR>
					<TD height="26">Documento:</TD>
					<TD height="26"><INPUT id="Text8" type="text" maxLength="30" size="30" name="Documento" DESIGNTIMEDRAGDROP="473"></TD>
				</TR>
				<TR>
					<TD height="26">Endereço:</TD>
					<TD height="26"><INPUT id="Text9" type="text" maxLength="50" size="50" name="Endereco" DESIGNTIMEDRAGDROP="474"></TD>
				</TR>
				<TR>
					<TD height="26">Bairro/Cidade:</TD>
					<TD height="26"><INPUT id="Text11" type="text" maxLength="50" size="50" name="BairroCidade" DESIGNTIMEDRAGDROP="475"></TD>
				</TR>
				<TR>
					<TD height="26">CEP/UF:</TD>
					<TD height="26"><INPUT id="Text12" type="text" maxLength="50" size="50" name="CepUF"></TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"></TD>
				</TR>
			</TABLE>
			<BR>
			<TABLE id="Table3" cellSpacing="1" cellPadding="1" width="90%" align="center" border="0">
				<TR>
					<TH colSpan="2">
					Informações sobre o Pagamento</TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"></TD>
				</TR>
				<TR>
					<TD>Valor Total:</TD>
					<TD><INPUT id="Text10" type="text" maxLength="10" size="10" name="Valor"></TD>
				</TR>
				<TR>
					<TD>Data de Vencimento:</TD>
					<TD><INPUT id=Text4 type=text maxLength=10 size=10 value="<%=FormatDateTime(Now+5,2)%>" name=Vencimento></TD>
				</TR>
				<TR>
					<TD vAlign="top">Demonstrativo:</TD>
					<TD><TEXTAREA id="Textarea1" name="Demonstrativo" rows="2" cols="40"></TEXTAREA></TD>
				</TR>
				<TR>
					<TD vAlign="top">Instruções:</TD>
					<TD><TEXTAREA id="Textarea2" name="Instrucoes" rows="2" cols="40">Não receber após o vencimento</TEXTAREA></TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"></TD>
				</TR>
			</TABLE>
			<P align="center">Enviar link do boleto para e-mail:<INPUT id="Text1" type="text" maxLength="100" size="30" name="email" value="seuemail@dominio.com.br"><BR>
				<br>
				<input type="submit" name="RUN" value="Gerar Boleto">
			</P>
		</form>
		<%
oConfig.Close()
oDB.Close()
%>
		<P align="center">[&nbsp; <A href="http://www.boletoasp.com">
				<b>Compre os códigos fontes</b></A> ]<BR>
			<BR>
			[ <A href="PartesBoleto.htm">Partes dos códigos</A>| <A href="FuncoesBoleto.asp">Algumas&nbsp;Funções</A>&nbsp;|
			<A href="../DOC">Documentação</A> ]<BR>
		</P>
	</body>
</html>
