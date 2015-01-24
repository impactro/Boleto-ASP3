<html>
  <head>
    <title>GeraBoleto</title>
  </head>
  <body>
<%

'Adapte conforme sua necessidade

'Força mudança de layout para formato brasileiro
Session.LCID = 1046 

if Request("ID") ="" then
	Response.Write("Informe o numero da cobrança, em ?ID=(numero)")
	Response.End 
end if
nCobranca=Request("ID")

'Obtem uma Conexão definida em sua aplicação geralmente pelo Global.ASA
SET oDB=Server.CreateObject( "ADODB.Connection" )
'oDB.Open Application("ConnectionString") 
oDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ProdCobranca.mdb")

'Abre uma tabela de configurações para obter os dados do cedente
SET oConfig=Server.CreateObject( "ADODB.RecordSet" )
oConfig.Open "Configuracoes", oDB, 2, 2

'Abre uma tabela de cobrança buscando a informação por um ID informado
SET oCobranca=Server.CreateObject( "ADODB.RecordSet" )
oCobranca.Open "SELECT * FROM Cobrancas WHERE NossoNumero=" & nCobranca, oDB, 2, 2

if oCobranca.EOF then
	Response.Write("Cobrança não existe")
	Response.End 
end if

'Preenche um formulário com os devidos dados

%>
	<form name=BOLETO method="post" action="BoletoExibe.asp">
		<input type=hidden name="Cedente" value="<%=oConfig("Cedente")%>" >
		<input type=hidden name="Banco" value="<%=oConfig("Banco")%>" >
		<input type=hidden name="Agencia" value="<%=oConfig("Agencia")%>" >
		<input type=hidden name="Conta" value="<%=oConfig("Conta")%>" >
		<input type=hidden name="Carteira" value="<%=oConfig("Carteira")%>" >
		<input type=hidden name="Convenio" value="<%=oConfig("Convenio")%>" >
		<input type=hidden name="Modalidade" value="<%=oConfig("Modalidade")%>">
		<input type=hidden name="Sacado" value="<%=oCobranca("Sacado")%>" >
		<input type=hidden name="SacadoDOC" value="<%=oCobranca("Documento")%>" >
		<input type=hidden name="Endereco1" value="<%=oCobranca("Endereco")%>">
		<input type=hidden name="Endereco2" value="<%=oCobranca("BairroCidade")%>" >
		<input type=hidden name="Endereco3" value="<%=oCobranca("CepUF")%>" >
		<input type=hidden name="NossoNumero" value="<%=oCobranca("NossoNumero")%>" >
		<input type=hidden name="NumeroDocumento" value="<%=oCobranca("NossoNumero")%>" >
		<input type=hidden name="DataDocumento" value="<%=FormatDateTime(Now,2)%>">
		<input type=hidden name="DataVencimento" value="<%=FormatDateTime(oCobranca("Vencimento"),2)%>">
		<input type=hidden name="Valor" value="<%=oCobranca("Valor")%>" >
		<input type=hidden name="Instrucoes" value="<%=oCobranca("Instrucoes")%>" >
		<input type=hidden name="Demonstrativo" value="<%=oCobranca("Demonstrativo")%>">
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

  </body>
</html>

