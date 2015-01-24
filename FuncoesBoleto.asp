<html>
	<head>
		<title>Funções do Boleto v1.8</title>
		<LINK rel="stylesheet" type="text/css" href="Boleto.css">
	</head>
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
		<%
'Autor: Fábio Ferreira de Souza - www.impactro.com.br / www.boletoasp.com.br
Session.LCID = 1046 

'Estas são as principais funções do gerador de boletos
'Com um pouco de trabalho é possivel, criar o gerador de boletos a partir destas rotinas
'Todas estas rotinas, existem também em outros sites gratuitamente e prontas para download também.
'O custo do gerador de boleto é é significativamente baixo para que você não
'precise juntar tudo, e fazer uma série de testes 
'Esse sistema já está bem testado, já foi impresso mais de 5.000 mil boletos de diversos bancos e carteiras 
'apenas do clientes direto da Impactro, em seus respectivos sitemas de e-commerce.
'A Impactro Informática, já está vendendo esse fonte deste agosto de 2004
'A versão em .Net é fechada (sem fontes) mas é mais facil de utilizar e personalizar


'Ignora erros para testes
ON ERROR RESUME NEXT
ImagePath = "Imagens/"

'== Rotina de calculo do Modulo 10 ==
FUNCTION CalculaModulo10( cTexto )
	DIM nMultiplicador,nPos,nRes,nTotal
	
	nMultiplicador=(len(cTexto) mod 2) 
	nMultiplicador=nMultiplicador+1
	nTotal=0
	
	for nPos=1 to len(cTexto)
		nRes= mid(cTexto, nPos, 1) * nMultiplicador
		if nRes>9 then
			nRes=int(nRes/10) + (nRes mod 10)
		end if
		nTotal=nTotal+nRes
		if nMultiplicador=2 then
			nMultiplicador=1
		else
			nMultiplicador=2
		end if
	next
	
	nTotal=(( 10-(nTotal mod 10)) mod 10 )
	CalculaModulo10=nTotal
	
END FUNCTION

DIM VTotal
'== Rotina de calculo do Modulo 11 ==
FUNCTION CalculaModulo11( cTexto, nBase )
	DIM nContador, nNumero, nTotal
	DIM mMultiplicador, nResto, nResultado
	DIM cCaracter

	nTotal=0
	nMultiplicador=2
	
	For nContador=Len(cTexto) to 1 step -1
		
		cCaracter=Mid( cTexto, nContador ,1)
		nNumero=Int( cCaracter ) * nMultiplicador
		'response.Write(nNumero & "<br>")
		nTotal=nTotal+nNumero
		
		nMultiplicador=nMultiplicador+1
		if nMultiplicador>nBase then
			nMultiplicador=2
		end if
		
	Next
    
    VTotal=nTotal

	nTotal=nTotal*10
	nResultado=nTotal mod 11
	
	if nResultado=0 or nResultado=10 then
		CalculaModulo11=1
	else
		CalculaModulo11=nResultado
	end if
	
END FUNCTION


'== Calcula o Fator de vencimento ==
FUNCTION CalcFatVencimento(dDtVenc)
	if dDtVenc=DateValue( "2001/01/01" ) then
		CalcFatVencimento=0
	else
		CalcFatVencimento=dDtVenc - DateValue( "1997/10/07" )
	end if
END FUNCTION

'== Rotina para Calculo do digito do Nosso numero da Nossa Caixa e Banespa ==
FUNCTION CalculaPesos(cValor,cPeso,lDig)
	DIM nContador,nV,nP,nD,nTotal,nResto
	nTotal=0
	
	For nContador=Len(cValor) to 1 step -1
		
		nV=Int( Mid( cValor, nContador ,1) )
		nP=Int( Mid( cPeso, nContador ,1) )
		
		nD=(nV*nP)
		
		if lDig then
			nD=nD MOD 10
		end if
		
		nTotal=nTotal + nD
		
	Next
	
	nResto=nTotal MOD 10
	if nResto = 0 then
		CalculaPesos=0
	else
		CalculaPesos=10-nResto
	end if
	
END FUNCTION


'== Gera a sequencia de codigo de barras ==
SUB GeraCodBarras( cCodBarras )
	DIM cOut,f,texto 
	DIM fi, f1, f2, i
	DIM BarCodes(100)
	
	BarCodes(0)="00110"
	BarCodes(1)="10001"
	BarCodes(2)="01001"
	BarCodes(3)="11000"
	BarCodes(4)="00101"
	BarCodes(5)="10100"
	BarCodes(6)="01100"
	BarCodes(7)="00011"
	BarCodes(8)="10010"
	BarCodes(9)="01010"
	
	For f1=9 to 0 step -1
		For f2=9 to 0 step -1
			fi=f1 * 10 + f2
			texto=""
			For i = 1 to 5
				texto=texto & Mid( BarCodes(f1), i, 1) & Mid( BarCodes(f2), i, 1)
			Next
			BarCodes(fi)=texto
		Next
	Next
	
	cOut="pfbfpfbf"
	texto=cCodBarras
	
	if (Len( texto ) MOD 2)<>0 then
		texto="0" & texto
	end if
		
	DO WHILE Len( texto ) > 0
		
		i=Int( Mid( texto, 1, 2 ) )
		texto=Mid( texto, 3 )
		f=BarCodes(i)
		
		For i=1 to 10 STEP 2

			if Mid( f ,i, 1) = "0" then
				cOut=cOut & "pf"
			else
				cOut=cOut & "pl"
			end if
                
			if Mid( f, i+1, 1) = "0" then
				cOut=cOut & "bf"
			else
				cOut=cOut & "bl"
			end if

		Next
		
	LOOP

	cOut=cOut & "pl"
	cOut=cOut & "bf"
	cOut=cOut & "pf"
	
	For i=1 to Len( cOut ) STEP 2
		SELECT CASE Mid( cOut, i, 2 )
			CASE "bf"
				Response.Write( "<img src='" + ImagePath + "b.gif' height=50 width=1>" )
			CASE "pf"
				Response.Write( "<img src='" + ImagePath + "p.gif' height=50 width=1>" )
			CASE "bl"
				Response.Write( "<img src='" + ImagePath + "b.gif' height=50 width=3>" )
			CASE "pl"
				Response.Write( "<img src='" + ImagePath + "p.gif' height=50 width=3>" )
		END SELECT
	Next

END SUB

%>
		<form id="Form1" method="post" action="FuncoesBoleto.asp">
			<br>
			<center><strong>Página de teste das principais rotinas existentes no Gerador de Boleto 
					ASP</strong></center>
			<br>
			<TABLE style="BACKGROUND-COLOR: #ffffcc" id="Table2" cellSpacing="1" cellPadding="1" width="100%"
				border="0">
				<TR>
					<TD vAlign="top" width="33%">
						<TABLE class="TableFormat" id="Table3" cellSpacing="1" cellPadding="1" width="80%" align="center"
							border="0" DESIGNTIMEDRAGDROP="13">
							<TR>
								<TH colSpan="2">
									Módulo 10
								</TH>
							</TR>
							<TR>
								<TD align="center" colSpan="2" height="10"></TD>
							</TR>
							<TR>
								<%
									cMOD10VALOR=Request("MOD10VALOR")
									if cMOD10VALOR="" then
										cMOD10VALOR="123456789"
									end if
									%>
								<TD width="50">Valor:</TD>
								<TD><INPUT id="Text2" type="text" maxLength="50" size="15" value="<%=cMOD10VALOR%>" name="MOD10VALOR" >&nbsp;
									<INPUT id="Submit2" type="submit" value="Testar" name="MOD10"></TD>
							</TR>
							<TR>
								<TD align="center" colSpan="2" height="50">
									<%
									Response.Write( CalculaModulo10( cMOD10VALOR ) )
									%>
								</TD>
							</TR>
						</TABLE>
						<br>
						<P>
							<TABLE class="TableFormat" id="Table1" cellSpacing="1" cellPadding="1" width="80%" align="center"
								border="0" DESIGNTIMEDRAGDROP="192">
								<TR>
									<TH colSpan="2">
										Módulo 11 (Padrão)
									</TH>
								</TR>
								<TR>
									<TD align="center" colSpan="2" height="10"></TD>
								</TR>
								<TR>
									<%
										cMOD11VALOR=Request("MOD11VALOR")
										if cMOD11VALOR="" then
											cMOD11VALOR="123456789"
										end if
										
										cMOD11BASE=Request("MOD11BASE")
										if cMOD11BASE="" then
											cMOD11BASE="9"
										end if
										
									%>
									<TD width="50">Valor:</TD>
									<TD><INPUT id=MOD11VALOR type=text maxLength=50 size=15 
            value="<%=cMOD11VALOR%>" name=MOD11VALOR>&nbsp;
									</TD>
								</TR>
								<TR>
									<TD width="50">Base:</TD>
									<TD><INPUT id=MOD11BASE type=text maxLength=1 size=1 
            value="<%=cMOD11BASE%>" name=MOD11BASE>&nbsp;&nbsp; <INPUT id="Submit1" type="submit" value="Testar" name="MOD11"></TD>
								</TR>
								<TR>
									<TD align="center" colSpan="2" height="50">
										<%
									Response.Write( "<b>" & CalculaModulo11( cMOD11VALOR, cInt(cMOD11BASE) ) & "</b> ( Total: " & VTotal & ")")
									%>
									</TD>
								</TR>
							</TABLE>
						</P>
					</TD>
					<TD vAlign="top" width="33%">
						<TABLE class="TableFormat" id="Table5" cellSpacing="1" cellPadding="1" width="80%" align="center"
							border="0">
							<TR>
								<TH colSpan="2">
									Fator de Vencimento</TH></TR>
							<TR>
								<TD align="center" colSpan="2" height="10"></TD>
							</TR>
							<TR>
								<%
									cDataValor=Request("DataValor")
									if cDataValor="" then
										cDataValor=FormatDateTime(Now,2)
									end if
									%>
								<TD width="50">Data:</TD>
								<TD><INPUT id=Text5 type=text maxLength=50 size=10 
            value="<%=cDataValor%>" name=DataValor>&nbsp; <INPUT id="Submit4" type="submit" value="Testar" name="TesteData"></TD>
							</TR>
							<TR>
								<TD align="center" colSpan="2" height="50">
									<%
									Response.Write( Right( "000" & CalcFatVencimento( DateValue( cDataValor ) ), 4 ) )
									%>
								</TD>
							</TR>
						</TABLE>
						<P>
							<TABLE class="TableFormat" id="Table6" cellSpacing="1" cellPadding="1" width="80%" align="center"
								border="0">
								<TR>
									<TH colSpan="2">
										Módulo&nbsp;Pesos
									</TH>
								</TR>
								<TR>
									<TD align="center" colSpan="2" height="10"></TD>
								</TR>
								<TR>
									<%
									cMODPesosVALOR=Request("MODPesosVALOR")
									cMODPesos=Request("MODPesos")
									if cMODPesosVALOR="" then
										cMODPesosVALOR="123456789"
										cMODPesos="731973197319"
									end if
									%>
									<TD width="50">Valor:</TD>
									<TD><INPUT id=Text1 type=text maxLength=50 size=15 value="<%=cMODPesosVALOR%>" name=MODPesosVALOR></TD>
								</TR>
								<TR>
									<TD>Pesos:</TD>
									<TD><INPUT id="Text3" type=text maxLength=50 size=15 value="<%=cMODPesos%>" name=MODPesos>&nbsp;
										<INPUT id="Submit6" type="submit" value="Testar" name="MODP"></TD>
								</TR>
								<TR>
									<TD align="center" colSpan="2" height="50">
										<%
									Response.Write( CalculaPesos( cMODPesosVALOR, cMODPesos,true ) )
									%>
									</TD>
								</TR>
							</TABLE>
						</P>
						<P>&nbsp;</P>
					</TD>
				</TR>
				<TR>
					<TD colSpan="2">
						<TABLE class="TableFormat" id="Table4" cellSpacing="1" cellPadding="1" width="90%" align="center"
							border="0">
							<TR>
								<TH colSpan="2">
									Código de Barras
									<TR>
										<TD align="center" colSpan="2" height="10"></TD>
									</TR>
									<TR>
										<%
									cCODBAR=Request("CODBARVALOR")
									if cCODBAR="" then
										cCODBAR="01234567890123456789012345678901234567890"
									end if
									%>
										<TD width="50">Valor:</TD>
										<TD><INPUT id="Text4" type="text" maxLength="50" size="50" value="<%=cCODBAR%>"
											name="CODBARVALOR" >&nbsp; <INPUT id="Submit3" type="submit" value="Testar" name="CODBAR"></TD>
									</TR>
									<TR>
										<TD align="center" colSpan="2" height="50">
											<%
									GeraCodBarras( cCODBAR )
									%>
										</TD>
									</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<br>
			<center>
				As funções acimas foram desenvolvidas apenas em ASP<br>
				<br>
				<br>
				[&nbsp;<A href="http://www.boletoasp.com/"><b>Compre os códigos 
						fontes</b></A> ]<BR>
				<BR>
				[ <A href="PartesBoleto.htm">Partes dos códigos</A>| <A href="FuncoesBoleto.asp">Algumas&nbsp;Funções</A>&nbsp;
				|&nbsp;<A href="../DOC">Documentação</A>
				]</center>
		</form>
	</body>
</html>
