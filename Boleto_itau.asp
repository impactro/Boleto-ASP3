<html>
<body>
<!-- 
ATEN��O: A m� configura��o, ou configura��o parcial das informa��es abaixo pode gerar um erro no programa
=======
ASP: Verifique se a configura��o do 'Session.LCID' (ASP) est� devidamente configurado para gerar datas no formato dd/mm/yyyy
Retire: "http://www.boletoasp.com.br" da linha abixo para gerar local 
-->
<form name=BOLETO method="post"action="BoletoExibe.asp" >
	
	<!-- Informa��es sobre o Cedente => Quem Emite o Boleto (Conta,Carteirar,etc) -->
	<input type=hidden name="Cedente" value="CEDENTE" >
	<input type=hidden name="Banco" value="341-7">
	<input type=hidden name="Agencia" value="6439" >
	<input type=hidden name="Conta" value="05531-7" >
	<input type=hidden name="Carteira" value="175">
	
	<!-- Informa��es sobre o Sacado Cedente => Quem tem que pagar o Boleto (Nome, Documento, endere�o) -->
	<input type=hidden name="Sacado" value="<%=Request("Sacador")%>">
	<input type=hidden name="SacadoDOC" value="">
	<input type=hidden name="Endereco1" value="<%=Request("Endereco")%>" >
	<input type=hidden name="Endereco2" value="<%=Request("Bairro") & " - " & Request("Cidade")%>" >
	<input type=hidden name="Endereco3" value="<%=Request("Estado") & " - CEP: " & Request("CEP")%>" >
	
	<!-- Informa��es sobre o Boleto => S�o os Valores e Datas e Numeros de controle para identificar o boleto (Valor, Data de Vencimento, Instru��es) -->
	<input type=hidden name="NossoNumero" value="<%=Request("NossoNumero")%>" >
	<input type=hidden name="NumeroDocumento" value="<%=Request("NumeroDoc")%>">
	<input type=hidden name="DataDocumento" value="<%=Request("DataDocumento")%>" >
	<input type=hidden name="DataVencimento" value="<%=Request("DataVencimento")%>">
	<input type=hidden name="Valor" value="<%=Request("ValorDocumento")%>" >
	<input type=hidden name="Instrucoes" value="Sr. caixa nao receberap�s vencimento." >
	<input type=hidden name="Demonstrativo" value="">
	<input type=hidden name="ImagePath" value="imagens/">
	
	<!-- Bot�o para submiss�o dos dados a p�gina BoletoExibe.asp que com os parametros acima gera o boleto (n�o � necess�rio, � apenas um bot�o para disparar o envio das vari�veis por POST) -->
	<input type=submit value="Testar" ID="Submit1" NAME="Submit1">
</form>

<script>
//o comando java script abaixo efetua uma auto-requisi��o:
//comente ou exclua as linha abaixo caso n�o queira utilizar uma auto submiss�o
onload=Start;

function Start(){
	
	document.BOLETO.submit();
}
</script>

</body>
</html>