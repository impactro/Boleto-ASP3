<html>
	<HEAD>
		<title>Formulario de Geração de Boleto</title>
<%
'A linha abaixo em Linux APACHE não dá suporte a LCID
'Session.LCID = 1046 'Força mudança de layout para formato brasileiro
'Response.Write( Session.LCID )
%>
		<script language="javascript">
<!--

var cBanco="001-9";

// Os valores predefinids abaixos são extraido da documentação fornecida de
// de cada banco, baixe os arquivos de: http://exemplos.boletoasp.com.br/DOC
function SelectBanco()
{
	var oOption;
	var n;
	if( cBanco=="001-9" || cBanco=="151-1" || cBanco=="399-9")
	{
		Formulario.Carteira.options.remove(0);
	}
		
	cBanco=Formulario.Banco.options[Formulario.Banco.selectedIndex].value;
	switch( cBanco )
	{
	
		case "001-9":	// Banco do Brasil

			Formulario.Carteira.options[0].value="16";
			Formulario.Carteira.options[0].text="Carteira 16"

			Formulario.Modalidade.options[0].value="18";
			Formulario.Modalidade.options[0].text="Barra 18"
			
			Formulario.Modalidade.options[1].value="19";
			Formulario.Modalidade.options[1].text="Barra 19"
			
			oOption=document.createElement("OPTION");
			oOption.value = "21";
			oOption.text = "Barra 21";
			Formulario.Modalidade.options.add(oOption);

			oOption=document.createElement("OPTION");
			oOption.text = "Carteira 18";
			oOption.value = "18";
			Formulario.Carteira.options.add(oOption);
			
			Formulario.Convenio.disabled=false;
			Formulario.Modalidade.disabled=false;
			Formulario.CodCedente.disabled=true;
			Formulario.Agencia.value="1234-5";
			Formulario.Conta.value="1234567-8";
			Formulario.Convenio.value="1234567";
			Formulario.NossoNumero.value="105";
			Formulario.DataVencimento.value="11/8/2004";
			Formulario.Valor.value="340,00";

			break;
			
		case "027-2":	// BESC

			Formulario.Carteira.options[0].value="25";
			Formulario.Carteira.options[0].text="Carteira Padrão"
			
			Formulario.Convenio.disabled=false;
			Formulario.Modalidade.disabled=true;
			Formulario.CodCedente.disabled=false;
			Formulario.Agencia.value="123";
			Formulario.Conta.value="12345678";
			Formulario.NossoNumero.value="7469108";

			break;
			
//		case "033-7":	// Banespa

//			Formulario.Carteira.options[0].value="01";
//			Formulario.Carteira.options[0].text="Carteira Padrão"
//			
//			Formulario.Convenio.disabled=true;
//			Formulario.Modalidade.disabled=true;
//			Formulario.CodCedente.disabled=false;
//			Formulario.CodCedente.value="40013012168";
//			Formulario.Agencia.value="123";
//			Formulario.Conta.value="12345678";
//			Formulario.NossoNumero.value="7469108";

//			break;

		case "041-8":	// Banrisul

			Formulario.Carteira.disabled=true;
			Formulario.Convenio.disabled=true;
			Formulario.Modalidade.disabled=true;
			Formulario.CodCedente.disabled=false;
			Formulario.Agencia.value="100.81";
			Formulario.Conta.value="12345.67";
			Formulario.CodCedente.value="0000001.83";
			Formulario.NossoNumero.value="22832563";
			Formulario.DataVencimento.value="4/7/2000";
			Formulario.Valor.value="550,00";
			
			break;
			
		case "104-0":	// Caixa Economica Federal

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira 01"
			
			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=true;
			Formulario.Modalidade.disabled=true;
			Formulario.Agencia.value="0238-2";
			Formulario.Conta.value="003376-1";

			break;		
			
		case "151-1":	// Nossa Caixa

			Formulario.Carteira.options[0].value="9";
			Formulario.Carteira.options[0].text="Carteira 9"
			
			Formulario.Modalidade.options[0].value="01";
			Formulario.Modalidade.options[0].text="01"

			Formulario.Modalidade.options[1].value="04";
			Formulario.Modalidade.options[1].text="04"
			
			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=true;
			Formulario.Modalidade.disabled=false;
			Formulario.Modalidade.selectedIndex=1;
			Formulario.Modalidade.value="04"
			Formulario.Agencia.value="0001-9";
			Formulario.Conta.value="002818-4";
			Formulario.Valor.value="350,00";
			Formulario.NossoNumero.value="990000001";
			Formulario.DataVencimento.value="15/07/2000";

			break;	
			
		case "237-2":	// Bradesco

			Formulario.Carteira.options[0].value="06";
			Formulario.Carteira.options[0].text="Carteira 06"
			
			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=true;
			Formulario.Modalidade.disabled=true;
			Formulario.Agencia.value="1234-5";
			Formulario.Conta.value="1234567-8";

			break;
			
		case "341-7":	// ITAU

			Formulario.Carteira.options[0].value="110";
			Formulario.Carteira.options[0].text="Carteira 110"

            oOption=document.createElement("OPTION");
			oOption.text = "Carteira 175";
			oOption.value = "175";
			Formulario.Carteira.options.add(oOption);
			
		    oOption=document.createElement("OPTION");
			oOption.text = "Carteira 198";
			oOption.value = "198";
			Formulario.Carteira.options.add(oOption);
			
			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=false;
			Formulario.Modalidade.disabled=true;
			Formulario.Agencia.value="0057";
			Formulario.Conta.value="12345-7";
			Formulario.Valor.value="123,45";
			Formulario.NossoNumero.value="12345678";
			Formulario.DataVencimento.value="01/05/2002";
			
			break;
			
		case "347-6":	// Sudameris

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira Sem Registro"

			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=true;
			Formulario.Modalidade.disabled=true;
			Formulario.Agencia.value="0501";
			Formulario.Conta.value="6703255-1";
			Formulario.Valor.value="35,00";
			Formulario.NossoNumero.value="0003020";
			Formulario.DataVencimento.value="02/10/2001";

			break;
			
		case "353-0": // Santander
		case "033-7": // Banespa

		    Formulario.Carteira.options[0].value="102";
			Formulario.Carteira.options[0].text="Carteira 102"
			Formulario.Convenio.disabled=true;
			Formulario.Modalidade.disabled=true;
			
			Formulario.Modalidade.options[0].value="4";
			Formulario.Modalidade.options[0].text="Identificador 4"

			Formulario.CodCedente.disabled=false;
			Formulario.Agencia.value="1234";
			Formulario.Conta.value="123456";
			Formulario.CodCedente.value="0282033";
			Formulario.Valor.value="273,71";
			Formulario.NossoNumero.value="5666124578000";
			Formulario.DataVencimento.value="15/5/2003";
		    
		    break;

		case "356-5":	// Real

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira Sem Registro"

			Formulario.Convenio.disabled=true;
			Formulario.CodCedente.disabled=true;
			Formulario.Modalidade.disabled=true;
			Formulario.Agencia.value="0501";
			Formulario.Conta.value="6703255-1";
			Formulario.Valor.value="35,00";
			Formulario.NossoNumero.value="0003020";
			Formulario.DataVencimento.value="02/10/2001";

			break;

		case "399-9":	// HSBC

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira Sem Registro"

			oOption=document.createElement("OPTION");
			oOption.text = "Carteira Com Registro";
			oOption.value = "02";
			Formulario.Carteira.options.add(oOption);
			
			Formulario.Convenio.disabled=false;
			Formulario.Modalidade.disabled=false;
			
			Formulario.Modalidade.options[0].value="4";
			Formulario.Modalidade.options[0].text="Identificador 4"

			Formulario.Modalidade.options[1].value="5";
			Formulario.Modalidade.options[1].text="Identificador 5"
			
			Formulario.Modalidade.options[2].value="99";
			Formulario.Modalidade.options[2].text="Com Registro"

			Formulario.CodCedente.disabled=false;
			Formulario.Agencia.value="1996";
			Formulario.Conta.value="41078-73";
			Formulario.Convenio.value="50950";
			Formulario.CodCedente.value="351202";
			Formulario.Valor.value="1200,00";
			Formulario.NossoNumero.value="39104766";
			Formulario.DataVencimento.value="04/07/2000";

			break;
			
		case "409-0":	// Unibanco

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira Sem Registro (01)"

			Formulario.Convenio.disabled=true;
			Formulario.Modalidade.disabled=true;
			
			Formulario.Modalidade.options[0].value="4";
			Formulario.Modalidade.options[0].text="Identificador 4"

			Formulario.CodCedente.disabled=false;
			Formulario.Agencia.value="1234";
			Formulario.Conta.value="123456";
			Formulario.CodCedente.value="1234561";
			Formulario.Valor.value="1000,00";
			Formulario.NossoNumero.value="11223344556677";
			Formulario.DataVencimento.value="31/12/2001";
			break;

		case "422-7":	// Safra

			Formulario.Carteira.options[0].value="01";
			Formulario.Carteira.options[0].text="Carteira Sem Registro (01)"

			Formulario.Convenio.disabled=true;
			Formulario.Modalidade.disabled=true;
			
			Formulario.CodCedente.disabled=true;
			Formulario.Agencia.value="00400";
			Formulario.Conta.value="00027824-7";
			Formulario.Valor.value="180,84";
			Formulario.NossoNumero.value="26173001";
			Formulario.DataVencimento.value="04/07/2000";
		}
}

//-->
		</script>
		<LINK href="Boleto.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body>
<script type="text/javascript">

    var _gaq = _gaq || [];
    _gaq.push(['_setAccount', 'UA-2170502-3']);
    _gaq.push(['_setDomainName', 'boletoasp.com.br']);
    _gaq.push(['_trackPageview']);

    (function () {
        var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
        ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
        var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
    })();

</script>
        <form id="Formulario" name="Formulario" action="BoletoExibe.asp" method="post">
			<TABLE id="Table1" cellSpacing="1" cellPadding="1" width="90%" align="center" border="0">
				<TR>
					<TH colSpan="2">
					Informações do Cedente</TD>
				</TR>
				<TR>
					<TD align="center" colSpan="2" height="10">&nbsp;&nbsp;&nbsp;</TD>
				</TR>
				<TR>
					<TD width="88">Cedente:</TD>
					<TD><INPUT id="Text1" type="text" maxLength="50" size="50" value="Sua empresa" name="Cedente"></TD>
				</TR>
				<TR>
					<TD width="88">Banco:</TD>
					<TD><SELECT id="Select1" onchange="SelectBanco();" name="Banco">
							<OPTION value="001-9" selected>001-BANCO DO BRASIL</OPTION>
							<OPTION value="027-2">027-BESC</OPTION>
							<OPTION value="033-7">033-BANESPA/SANTANDER</OPTION>
							<OPTION value="041-8">041-BANRISUL</OPTION>
							<OPTION value="104-0">104-CAIXA</OPTION>
							<OPTION value="151-1">151-NOSSA CAIXA</OPTION>
							<OPTION value="237-2">237-BRADESCO</OPTION>
							<OPTION value="341-7">341-ITAU</OPTION>
							<OPTION value="347-6">347-SUDAMERIS</OPTION>
							<OPTION value="409-0">409-UNIBANCO</OPTION>
							<OPTION value="347-6">347-SUDAMERIS</OPTION>
							<OPTION value="353-0">353-SANTANDER</OPTION>
							<OPTION value="356-5">356-REAL</OPTION>
							<OPTION value="399-9">399-HSBC</OPTION>
							<OPTION value="409-0">409-UNIBANCO</OPTION>
							<OPTION value="422-7">422-SAFRA</OPTION>
						</SELECT></TD>
				</TR>
				<TR>
					<TD width="88">Agência:</TD>
					<TD><INPUT id="Text2" type="text" maxLength="6" size="6" value="1234-5" name="Agencia">&nbsp;(incluindo 
						os zeros a frente)</TD>
				</TR>
				<TR>
					<TD width="88">Conta:</TD>
					<TD><INPUT id="Text3" type="text" maxLength="9" size="9" value="1234567-8" name="Conta">&nbsp;(incluindo 
						os zeros a frente)</TD>
				</TR>
				<TR>
					<TD width="88">Carteira</TD>
					<TD><SELECT id="Select2" name="Carteira" DESIGNTIMEDRAGDROP="63">
							<OPTION value="16" selected>Carteira 16</OPTION>
							<OPTION value="18">Carteira 18</OPTION>
						</SELECT></TD>
				</TR>
				<TR>
					<TD width="88">Modalidade</TD>
					<TD><SELECT id="Select3" name="Modalidade">
							<OPTION value="19" selected>Barra 19</OPTION>
							<OPTION value="21">Barra 21</OPTION>
						</SELECT></TD>
				</TR>
				<TR>
					<TD width="88">Convénio</TD>
					<TD><INPUT id="Text13" type="text" maxLength="7" size="7" value="123456" name="Convenio"></TD>
				</TR>
				<TR>
					<TD width="88">Cod.Cedente</TD>
					<TD><INPUT id="Text14" disabled type="text" maxLength="11" size="11" value="12345" name="CodCedente"></TD>
				</TR>
				<TR>
					<TD width="88" colSpan="2" height="10"></TD>
				</TR>
			</TABLE>
			<br>
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
					<TD height="26"><INPUT id="Text5" type="text" maxLength="50" size="50" value="Nome do cliente/Razão social"
							name="Sacado" DESIGNTIMEDRAGDROP="421"></TD>
				</TR>
				<TR>
					<TD height="26">Documento:</TD>
					<TD height="26"><INPUT id="Text8" type="text" maxLength="30" size="30" value="RG: 123.321.123-8" name="SacadoDOC"
							DESIGNTIMEDRAGDROP="473"></TD>
				</TR>
				<TR>
					<TD height="26">Endereço:</TD>
					<TD height="26"><INPUT id="Text9" type="text" maxLength="50" size="50" value="Rua fulano de Tal, 123 Ap 123"
							name="Endereco1" DESIGNTIMEDRAGDROP="474"></TD>
				</TR>
				<TR>
					<TD height="26">Bairro/Cidade:</TD>
					<TD height="26"><INPUT id="Text11" type="text" maxLength="50" size="50" value="Bairro - Cidade" name="Endereco2"
							DESIGNTIMEDRAGDROP="475"></TD>
				</TR>
				<TR>
					<TD height="26">CEP/UF:</TD>
					<TD height="26"><INPUT id="Text12" type="text" maxLength="50" size="50" value="CEP: 12345-678 SP" name="Endereco3"></TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"></TD>
				</TR>
			</TABLE>
			<br>
			<TABLE id="Table3" cellSpacing="1" cellPadding="1" width="90%" align="center" border="0">
				<TR>
					<TH colSpan="2">
					Informações sobre o Pagamento</TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"><INPUT type="hidden" value="Imagens/" name="ImagePath"></TD>
				</TR>
				<TR>
					<TD>Numero do documento:</TD>
					<TD><INPUT id="Text6" type="text" maxLength="10" size="10" value="23456" name="NumeroDocumento"></TD>
				</TR>
				<TR>
					<TD>Nosso Numero:</TD>
					<TD><INPUT id="Text7" type="text" maxLength="17" size="15" value="12345678" name="NossoNumero"></TD>
				</TR>
				<TR>
					<TD>Valor Total:</TD>
					<TD><INPUT id="Text10" type="text" maxLength="10" size="10" value="1234,56" name="Valor"></TD>
				</TR>
				<TR>
					<TD>Data do Documento:</TD>
					<TD><INPUT id=Text17 type=text maxLength=10 size=10 
      value="<%=FormatDateTime(Now,2)%>" name=DataDocumento 
    ></TD>
				</TR>
				<TR>
					<TD>Data de Vencimento:</TD>
					<TD><INPUT id=Text4 type=text maxLength=10 size=10 
      value="<%=FormatDateTime(Now,2)%>" name=DataVencimento 
      >&nbsp;(informe 01/01/2001 para contra apresentação)</TD>
				</TR>
				<TR>
					<TD vAlign="top">Demonstrativo:</TD>
					<TD><TEXTAREA id="Textarea1" name="Demonstrativo" rows="2" cols="40">Referente serviços de manutenção</TEXTAREA></TD>
				</TR>
				<TR>
					<TD vAlign="top">Instruções:</TD>
					<TD><TEXTAREA id="Textarea2" name="Instrucoes" rows="2" cols="40">Não receber após o vencimento
						</TEXTAREA></TD>
				</TR>
				<TR>
					<TD colSpan="2" height="10"></TD>
				</TR>
			</TABLE>
			<P align="center">Clique no botão abaixo para ver como será o boleto gerado<br>
				<br>
				<INPUT id="Submit1" type="submit" value="Gerar Boleto de Teste" name="Submit1"></P>
			<br>
			<P align="center">[&nbsp; <A href="http://www.boletoasp.com/">
					<b>Compre os códigos fontes</b></A> ]<BR>
				<BR>
				[ <A href="PartesBoleto.htm">Partes dos códigos</A>|&nbsp; <A href="FuncoesBoleto.asp">
					Algumas&nbsp;Funções</A>&nbsp;|&nbsp; <A href="../DOC">Documentação</A> ]<BR>
			</P>
		</form>
	</body>
</html>
