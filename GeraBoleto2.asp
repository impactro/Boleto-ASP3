<html>
  <head>
    <title>GeraBoleto 2</title>
  </head>
  <body>
<%

'Exemplo baseado no programa Geraboleto.asp que executa a leitura de uma sistema qualquer.
'Adapte conforme sua necessidade

Session.LCID = 1046 'Força mudança de layout para formato brasileiro
'Response.Write( Session.LCID )

cCobranca=Request("NumDOC")
if cCobranca<>"" then
	cNome=Request("Nome")
	cEndereco=Request("Endereco")
	cBairro=Request("Bairro")
	cCidade=Request("Cidade")
	cEstado=Request("Estado")
	cCEP=Request("CEP")
	nTotal=Request("Total")
	dtVenc=cDate(Request("DtVenc"))
	cDemostrativo=Request("Dem")

	'Preenche um formulário com os devidos dados, e posta ao programa padrão
	
	'Apos comprar o boleto, retire as referencias "http://www.boletoasp.com.br"
%>
	<form name=BOLETO method="post" action="http://www.boletoasp.com.br/BoletoExibe.asp">
		<input type=hidden name="Cedente" value="Sua Empresa" >
		<input type=hidden name="Banco" value="001-9" >
		<input type=hidden name="Agencia" value="9999" >
		<input type=hidden name="Conta" value="9999999-9" >
		<input type=hidden name="Carteira" value="18" >
		<input type=hidden name="Modalidade" value="18">
		<input type=hidden name="Sacado" value="<%=cNome%>" >
		<input type=hidden name="SacadoDOC" value="" >
		<input type=hidden name="Endereco1" value="<%=cEndereco%>">
		<input type=hidden name="Endereco2" value="<%=cBairro & " - " & cCidade%>" >
		<input type=hidden name="Endereco3" value="<%="CEP: " & cCEP & " " & cEstado%>" >
		<input type=hidden name="NossoNumero" value="<%=cCobranca%>" >
		<input type=hidden name="NumeroDocumento" value="<%=cCobranca%>" >
		<input type=hidden name="DataDocumento" value="<%=FormatDateTime(Now,2)%>">
		<input type=hidden name="DataVencimento" value="<%=FormatDateTime(dtVenc,2)%>">
		<input type=hidden name="Valor" value="<%=nTotal%>" >
		<input type=hidden name="Instrucoes"	value="Sr. caixa nao receberapós vencimento.<br><br>Após vencimento a compra é cancela automaticamente." >
		<input type=hidden name="Demonstrativo" value="<%=cDemostrativo%>">
		<input type=hidden name="ImagePath" value="imagens/">
		<input type=submit name="RUN" value="Gerar Boleto">
	</form>

<script>
// Efetua um auto submit
onload=Start;

function Start()
{
	document.BOLETO.submit();
}
</script>
<%
else
	'Esta URL Geralmente pode funcionar, mas é facilmente burlada e pode gerar erro no programa se um dos parametros contiver caracteres especiais como '&' ou '%'
	'cURL="GeraBoleto2.asp?NumDOC=12&Nome=Fábio Ferreira de Souza&Endereco=Rua Esquina da Curva&Bairro=Parque Fim do Mundo&Cidade=Sei La Aonde&Estado=SP&CEP=03014-000&Total=1.234,56&DtVenc=26/06/2005"
	
	'O exemplo abaixo é mais complexo porem já trata a URL com caracteres especiais, utilizando a função "Escape"
	cNome="Fabio & Erika"
	cEndereco="XPTO!@#@%#%$^%$@$ Endereço complexo"
	cBairro="Parque Fim do Mundo"
	cCidade="Sei La Aonde"
	
	cURL="GeraBoleto2.asp?NumDOC=12&Nome=" & Escape( cNome ) & _
		"&Endereco="& Escape( cEndereco ) & _
		"&Bairro=" & Escape( cBairro ) & _
		"&Cidade=" & Escape( cCidade ) & _
		"&Dem=" & Escape( "Referente a compra através do site: www.seusite.com.br") & _
		"&Estado=SP&CEP=03014-000&Total=1.234,56&DtVenc=26/06/2005"

%>
	Exemplo do link existente na web ou no e-mail<br>
	Clique <a href="<%=cURL%>">aqui</a> para visualizar o boleto
	<br>
	<br>
	<br>
	<br>
	<br>
	<br>
				[ <A href="http://www.boletoasp.com.br">Home</A> | <A href="http://www.boletoasp.com.br/PartesBoleto.htm">Partes dos códigos</A> | <A href="http://www.boletoasp.com.br/FuncoesBoleto.asp">Algumas Funções</A> |
				 <A href="http://exemplos.boletoasp.com.br/DOC">Documentação</A> ]<BR>
<%
end if
%>
  </body>
</html>

