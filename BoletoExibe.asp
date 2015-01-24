<html>
<head>
    <title>Boleto</title>
    <style type="text/css">
        BODY { FONT-SIZE: 10pt; FONT-FAMILY: Arial; BACKGROUND-COLOR: #ffffff }
        TABLE { BORDER-LEFT: #000000 1px solid; BORDER-BOTTOM: #000000 1px solid }
        TD { BORDER-RIGHT: #000000 1px solid; BORDER-TOP: #000000 1px solid; FONT-SIZE: 7pt; FONT-FAMILY: Arial }
        .noborder { BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px }
        .campo { FONT-WEIGHT: bold; FONT-SIZE: 8pt; MARGIN-LEFT: 2px; MARGIN-RIGHT: 2px; FONT-FAMILY: Verdana }
    </style>
    <!-- (não remova) BoletoExibe 2.9.9.18 - 18/09/2009 (não remova) -->
</head>
<body>

<%
' Copyright (C) 2004-2008 - IMPACTRO Informática - Fábio Ferreira de Souza
' Versão 2.9.9.18
' Desenvolvimento inicial em C#.Net -> adaptado para ASP em 4/9/2004
' Suporte para Linux Apache Sun ASP em 06/06/2006
' Suporte: fabio@impactro.com.br
' Vendas: http://www.boletoasp.com.br

'========================
' Bancos e Carteiras suportadas: (apenas cobrança sem registro)
' Banco do Brasil 001-9 - Várias Carteiras e Modalidades - Revisado de acordo com ultimas alteraçòes em 06/02/2006
' BESC            027-2 - Carteira 25
' Banespa		  033-7 - Carteira padrão - Ultima alteração: 27/03/2006 => Substituido pelo Santander em 22/12/2011
' Banrisul        041-8 - Carteira padrão
' Caixa	Econica   104-0 - Carteira padrão
' Nossa Caixa     151-1 - Carteira 9 - Modalidades: 01, 04, 09, 13, 16, 17, 18
' Bradesco        237-2 - Carteira padrão
' ITAU            341-7 - Carteira 175
' Sudameris		  347-6 - Carteira padrão
' Santander	      353-0 - Carteira 102
' Real            356-5 - Carteira padrão
' Banco Mercantil 389-1 - Carteira padrão
' HSBC            399-9 - Carteira 01, 02 - Modalidades: 4 e 5 (Carteira 02 é com registro)
' UNIBANCO        409-0 - Carteira padrão
' CITIBANK        745-5 - Carteira padrão
' SICOOB          756-0 - Carteira Padrão
'========================

' Força mudança de layout para formato brasileiro
' Caso seu sistema utilize outro formato LCID, reveja os tratamentos númericos e datas.
' Alguns servidores APACHE (Linux) que dão suporte a ASP nem sempre aceitam configurações LCID
'Session.LCID = 1046 
DebugResponse( "Session.LCID: " & Session.LCID )

'== Declaração e descrição das váriáveis

'== INFORMAÇÕES SOBRE O CEDENTE ==
Dim Banco			'Nº do Banco
Dim BancoDig		'Nº da Dígito do Banco 
Dim Agencia			'Nº da Agencia
Dim AgenciaDIG		'Nº da Dígito da Agencia
Dim Conta			'Nº da Conta
Dim ContaDIG		'Nº do Dígito da Conta
Dim Convenio		'Nº do convenio firmado com o banco do Brasil(Carteiras 171,172), ou Código da Coperativa para banco SICOOB
Dim CodCedente		'Nº do código do Sedente (HSBC) / Código do cliente (Unibanco)
Dim Carteira		'Nº da carteira 
Dim Modalidade		'Código da Modalidade (Tipo/Barra de formação do código de Barras) 
Dim Cedente			'Nome da empresa cedente emissora do boleto
DIM cAgenciaContaExibicao

'== INFORMAÇÔES DO SACADO ==
Dim Sacado			'Nome ou Razão social do sacado
Dim SacadoDoc		'Documento de identificação do sacado (CPF/CNPJ/RG)
Dim Endereco1		'Linha 1 do endereço (ex: Rua XXXX, Numero - Complemento)
Dim Endereco2		'Linha 2 do endereço (ex: Bairro - Cidade)
Dim Endereco3		'Linha 3 do Endereço (ex: CEP:12345-678 - UF)

'== INFORMAÇÔES DO BOLETO ==
Dim DataDocumento	'Data de emissão do documento
Dim DataVencimento	'Data de vencimento do documento
Dim NossoNumero		'Nosso Número
Dim NumeroDocumento	'Numero do documento
Dim Quantidade		'Valor do campo quantidade
Dim Valor			'Valor do documento
Dim Instrucoes		'Instruções para o caixa
Dim Demonstrativo	'Demonstrativo do conteudo do boleto
Dim ImagePath		'Caminho relativo onde estão as imagens
Dim Especial        'Variável logica para indicar quando a rotina Modulo 11 deve ser usada em logica especial

'Por padrão e compatibilidade, a logica default, e a normal (não especial)
Especial=false      'ao trocar para Especial (true) apos chamar a rotina de Modulo11 voltar para false

'== Inicializa variáveis
'Recomendamos que os dados sejam passados por meio de um POST
'veja o arquivo GeraBoleto.ASP que obtem as informaçãoes de um banco dedos
'e gera um formulário para ser postado

Cedente=Request("Cedente")

Banco=Request("Banco")
BancoDig=Mid( Banco, 5)
Banco=Mid(Banco,1,3)

Agencia=Request("Agencia")
Conta=Request("Conta")

Carteira=Request("Carteira")
Convenio=Request("Convenio")
CodCedente=Request("CodCedente")
Modalidade=Request("Modalidade")

Sacado=Request("Sacado")
SacadoDOC=Request("SacadoDOC")
Endereco1=Request("Endereco1")
Endereco2=Request("Endereco2")
Endereco3=Request("Endereco3")

NossoNumero=Request("NossoNumero")
NumeroDocumento=Request("NumeroDocumento")

'Em servidores APACHE ou IIS que não se pode configurar o LCID, as linhas abaixo de requisição
'e depois de exibição, aparte HTML poderão dar problemas, veja com seu provedor de hospedagem
'As funções GetDataVar() e FormatData() foram criadas para resolver esse problema.

if Request("DataDocumento")="" then
	DataDocumento=Now()
else
	DataDocumento=GetDataVar( "DataDocumento" )
end if

if Request("DataVencimento")="" then
	DataVencimento=DateValue( "2001/01/01" ) 'Sem data de vencimento
else
	DataVencimento=GetDataVar( "DataVencimento" )
end if

'Tem que ser formato brasileiro (força) para garantir compatibilidade sem o LCID e em Linux Apache
Valor=Trim(Request("Valor"))
Instrucoes=Request("Instrucoes")
Demonstrativo=Request("Demonstrativo")

ImagePath=Request("ImagePath")
if ImagePath="" then
	ImagePath="Imagens/"
end if

'== INFORMAÇÔES CALCULADAS ==
Dim LinhaDigitavel
Dim CodBarras

'== Inicializa o calculo das informações

CalculaCodigoBarras()
CalculaLinhaDigitavel()

'== Rotina que Gera o Codigo de Barras ==
SUB CalculaCodigoBarras()
	DIM cBanco,cCalcFat,cValor,cConvenio
	DIM cNossoNumero,cAgenciaNumero,cContaNumero,cCarteira,cModalidade,cCodCedente
	DIM cCodePadrao,cLivre,nDig

'	04. LEIAUTE DO CÓDIGO DE BARRAS PADRÃO (vale para qualquer banco)
'...............................................................    
'   N.    POSIÇÕES     PICTURE     USAGE        CONTEÚO                
'...............................................................    
'    01    001 a 003    9/003/      Display      Identificação do banco
'    02    004 a 004    9/001/      Display      9 /Real/
'(a) 03    005 a 005    9/001/      Display      DV /*/
'(b) 04    006 a 009    9/004/      Display      fator de vencimento
'    05    010 a 019    9/008/v99   Display      Valor
'    06    020 a 044    9/025/      Display      CAMPO LIVRE
'...............................................................    
'OBS: 1 - o digito verificador da 5 (quinta) posição é calculado     
'         com base no módulo 11 específico previsto no item 12;         
'OBS: 2 - o fator de vencimento é calculado com base na metodologia 
'         descrita no item 05.                                          

	cBanco=Right( "000" & Banco, 3 )
	cCalcFat=Right( "00" & CalcFatVencimento(DataVencimento), 4 )

	'Ajusta o valor para truncar e transformar em numeros sem ponto com 2 casas de precisão
	cValor=Replace(Valor,".","")
	DebugResponse("Valor: " & Valor )
	if InStr(cValor,",")>0 then
	    cValor=cValor & "00"
	    n=InStr(cValor,",")
	    cValor=mid(cValor,1,n-1) & mid( cValor, n+1, 2)
	else
	    cValor=cValor & "00"
	end if
	Valor=CDbl(cValor)/100
	cCodePadrao=cBanco & "9" & cCalcFat & Right("0000000000" & cValor,10)
	DebugResponse("Vencimento: " & cCalcFat & " (" & FormatData(DataVencimento) & ")" )
	
'O Modo de geração do Boleto bancário é padrão para qualquer banco.
'O Campo Livre muda de de banco para banco e de acordo com a carteira e convenios contratados
'Esta área pode e deve ser alterada conforme a necessidade
'Existem alguns DebugResponse() para auxiliar uma possível depuração
'Consulte a rotina DebugResponse e tire o comentário da linha Response.Write para exibir todo o rastreamento de depuração
'Sugiro que realise testes imprimindo alguns boletos e entregando a seu gerente
'para que possa ser validado o boleto e/ou  pagando pequenos valores com o boleto
'gerado antes de colocar em produção

'"cLivre" deve ser gerada de acordo com o banco para que o boleto continue a ser gerado
'deve ser obrigatoriamente uma string com 25 posições numericas

'A Diferença entre o IPTE e o CÓDIGO DE BARRAS
'IPTE: seqüência de 47 números que aparecerá na parte superior da ficha de compensação, que pode ser digitada como alternativa à leitura óptica do CÓDIGO DE BARRAS
'CÓDIGO DE BARRAS: seqüência de 44 números que aparecerá na parte inferior da ficha de compensação, porém escrita de forma codificada em barras
'Apesar de ambos serem uma seqüência de números, as regras que darão origem a esses códigos são diferentes.

	Banco=Right( "00" & Banco, 3 )

'Calculo do campo Livre 25 Posições
	SELECT CASE Banco

		CASE "001"	' Banco do Basil
			BancoDig="9"
			
			ContaDig=Mid(Conta, 9)
			Conta=Mid(Conta,1,7)
			
			cModalidade=Right( "0" & Modalidade, 2 )
			
			'Retira pontos traços etc, do nosso número
			cNossoNumero=Replace( NossoNumero,".","")
			cNossoNumero=Replace(cNossoNumero,"-","")
			cNossoNumero=Replace(cNossoNumero,"X","")
			NossoNumero=cNossoNumero
			
			if Len(Convenio) = 7 then
				'CÓDIGO DE BARRAS PARA EMISSÃO DE BLOQUETOS NAS CARTEIRAS 17 E 18,
				'EXCLUSIVO PARA CONVÊNIOS COM NUMERAÇÃO SUPERIOR À 1.000.000 (UM MILHÃO).
				cNossoNumero=Convenio & Right( "0000000000" & cNossoNumero, 10 )
				cLivre="000000" & _
						cNossoNumero & _
						Right( "0" & Carteira,2 )
						
				DebugResponse( "Convenio 7 Digitos: " & cLivre )
				
			elseif Len(Convenio) = 4 then
				
				cLivre= Right( "0000000" & Convenio,4) & _ 
				        Right( "0000000" & cNossoNumero,7) & _ 
						Right( "0000000" & mid(Agencia,1,4),4) & _
						Right( "0000000" & mid(Conta,1,8),8) & _
						Right( "00" & Carteira,2)
						
            elseif Len(Convenio) = 6 then
				
			    cNossoNumero=Convenio & Right( "00000" & cNossoNumero, 5 )
				
				NossoNumero=cNossoNumero & "-" & CalculaModulo11(cNossoNumero,9)
				
				if Carteira="16" or Carteira="18" then
					cNossoNumero=Right( "00000000000000000" & cNossoNumero, 17)
					if cModalidade="21" then
						cLivre=Convenio & _
							   cNossoNumero & _
							   cModalidade
					else 'Carteira 12 - Tipos de Convênio 2, 3, 4 ou 5. (COD CEDENTE=CONVENIO NO BB)
						cLivre=Right( "0000000000" & cNossoNumero,11) & _ 
							Right( "0000" & cAgenciaNumero,4) & _
							Right( "00000000" & Convenio,8) & _
							Right( "00" & cCarteira,2)
					end if
				else
					Response.Write("Carteira invalida")
					Response.End()
				end if
				
				DebugResponse( "Convenio 6 Digitos: " & cLivre )

        	else
        	
				Response.Write("Numero de convenio invalido")
				Response.End()
				
			end if
			
		CASE "027"  ' BESC
			BancoDig="2"

			cConvenio=Right( "00000" & Convenio, 5 )
			cCarteira=Right( "0" & Carteira, 2 )
			cNossoNumero=Right( "000000000000" & NossoNumero, 13 )

			cLivre=cConvenio & Mid(cNossoNumero,1,3) & cCarteira & Mid(cNossoNumero,4) & "027"

			cLivre=cLivre & CalculaModulo10( cLivre )
			cLivre=cLivre & CalculaModulo11( cLivre, 7 )

'		CASE "033"  ' Banespa
'			BancoDig="7"
'
'			cAgencia=right("000" & Split(Agencia,".")(0), 3) 'Ignora qualquer dígito e deixa com 3 posições
'			cCodCedente=Right( "00000000000" & Replace(CodCedente,"-",""), 11 ) 'Tira o traço do digito e deixa com 11 posições
'			cNossoNumero=Right( "0000000" & NossoNumero, 7 )
'			
'			cLivre=cCodCedente & cNossoNumero & "00" & "033"
'           nD1=CalculaModulo10(cLivre)
'			nD2=CalculaModulo11(cLivre & nD1,7)
'			do while nD2=-1 'Veja a rotina do modulo11
'				nD1=nD1+1
'   			nD2=CalculaModulo11(cLivre & nD1,7)	
'			loop
'			cLivre = cLivre & nD1 & nD2
'			
'			'Para gerar o digito do nosso numero
'			NossoNumero=cAgencia & " " & cNossoNumero & "-" & CalculaPesos(cAgencia & cNossoNumero,"7319731973",true)
'			
'			'Coloca o numero do documento igual ao nosso numero
'			NumeroDocumento=NossoNumero-->

    	CASE "041"	' Banco Banrisul
			BancoDig="8"
		
			cAgencia=Split(Agencia,".")(0)
			cCodCedente=Split(CodCedente,".")(0)
			cNossoNumero=Split(NossoNumero,".")(0)
			
			cAgenciaNumero=Right( "000" & cAgencia, 3 )
			cCodCedente=Right( "0000000" & cCodCedente, 7 )
			cNossoNumero=Right( "00000000" & cNossoNumero, 8 )

			cLivre="21" & _
				cAgenciaNumero & _
				cCodCedente & _
				cNossoNumero & _
				"041"
		
			cDV=CalculaModulo10( cLivre ) & CalculaModulo11( cLivre, 7 )
			cLivre=cLivre & cDV
			
			cDAC=CalculaModulo10(cNossoNumero)
			cDAC=cDAC & CalculaModulo11(cNossoNumero & cDAC, 7 )
			NossoNumero=NossoNumero & "-" & cDAC
			
		CASE "104": ' Caixa Economica Federal
		
			BancoDig="0"
	
			AgenciaDig=Mid(Agencia,6)
			Agencia=Mid(Agencia,1,4)
			
			if len(CodCedente)=15 then
			
			    cNossoNumero = Right("0000000000" & NossoNumero, 10)
                cLivre = cNossoNumero & CodCedente

                Especial=true
                NossoNumero = cNossoNumero & "-"  & CalculaModulo11(cNossoNumero, 9)
                Especial=false       
                
                cCodCedente = CodCedente & "." & CalculaModulo11(CodCedente, 9)
                cAgenciaContaExibicao = _
                        mid(cCodCedente, 1, 4) + "." & _
                        mid(cCodCedente, 5, 3) + "." & _
                        mid(cCodCedente, 8, 8) + "." & _
                        mid(cCodCedente, 17 , 1)

			else
			
			    ContaDig=Mid( Conta, 8)
			    Conta=Mid(Conta,1,6)

			    cAgenciaNumero=Right( "0000" & Agencia, 4 )
			    cContaNumero=Right( "000000" & CodCedente, 6 )
			    cCarteira=Right( "00" & Carteira, 2 )
			    cNossoNumero="9" & Right( "00000000000000000" & NossoNumero, 17 )

			    cLivre="1" & cContaNumero & cNossoNumero
			    NossoNumero=cNossoNumero & "-" & CalculaModulo11(cNossoNumero,9)
    			
			    'Verificar se há a necessidade
			    Conta=cContaNumero & "-" & CalculaModulo11(cAgenciaNumero & cContaNumero,9)
			
			end if

		CASE "151": ' Nossa Caixa
			BancoDig="1"
			
			Agencia=Mid(Agencia,1,4)
			ContaDig=Mid(Conta,8,1)
			Conta=Mid(Conta,1,6)
			Carteira=Modalidade
			cModalidade=Mid(Modalidade,2) 'pega o segundo digito da modalidade
			cNossoNumero=Mid(NossoNumero,1,9)

			cAgenciaNumero=Right( "0000" & Agencia, 4 )
			cContaNumero=Right( "000000" & Conta, 6 )
			cNossoNumero="99" & Right( "0000000" & cNossoNumero, 7 )

			'O campo nosso numero tem que iniciar com 9 e ter 9 digitos
			cLivre=cNossoNumero & _
				cAgenciaNumero & _
				cModalidade & _
				cContaNumero & _
				"151"
				
			nD1=CalculaModulo10(cLivre)
			nD2=CalculaModulo11(cLivre & nD1,7)
			do while nD2=-1 'Veja a rotina do modulo11
				nD1=nD1+1
    			nD2=CalculaModulo11(cLivre & nD1,7)	
			loop
			cLivre = cLivre & nD1 & nD2
			
			'Calucla o digito do Nosso Numero
			cModalidade=Right("0" & Modalidade,2) 'força o truncamento para ter sempre 2 digitos a modalidade
			'nTotal=DVNossaCaixa( cAgenciaNumero & cModalidade & "0" & cContaNumero ) + int(ContaDig)
			'nTotal=nTotal + DVNossaCaixa( cNossoNumero )
			'DebugResponse( "Total: " & nTotal )
			
			cDV=CalculaPesos(cAgenciaNumero & cModalidade & "0" & cContaNumero & ContaDig & cNossoNumero, "31973197319731319731973", false )
			
			nResto=nTotal MOD 10
			nResto=10-nResto

			NossoNumero=NossoNumero & "-" & cDV

		CASE "237": ' Bradesco
			BancoDig="2"
	
			AgenciaDig=Mid(Agencia,6)
			Agencia=Mid(Agencia,1,4)
			
			ContaDig=Mid(Conta, 9)
			Conta=Mid(Conta,1,7)
								
			cAgenciaNumero=Right( "0000" & Agencia, 4 )
			DebugResponse("Bradesco-Agencia: '" & cAgenciaNumero & "'")
			
			cCarteira=Right( "00" & Carteira, 2 )
			DebugResponse("Bradesco-Carteira: '" & cCarteira & "'")
			
			cNossoNumero=Right( "000000000" & NossoNumero, 9 )
			DebugResponse("Bradesco-NossoNumero: '" & cNossoNumero & "'")
			
			cContaNumero=Right( "0000000" & Conta, 7 )
			DebugResponse("Bradesco-Conta: '" & cContaNumero & "'")
			
			cModalidade=Right( "00" & Modalidade, 2 )
            DebugResponse("Bradesco-Modalidade: '" & cModalidade & "'")

			'Em alguns casos a constante "00" pode ser "11" é o tal "Ano do Nosso Número"
			'Para que isso fique parametrizavel este valor foi incluido como modalidade, 
			'que caso não seja especificada fica sendo 00
			cLivre=cAgenciaNumero & _
				cCarteira & _
				cModalidade & _
				cNossoNumero & _
				cContaNumero & "0"
				
            DebugResponse("Bradesco-CampoLivre: " & cLivre)

            Especial=true
			'para o digito verificador o nosso numero tem 11 digitos (Modalidade+NossoNumero)
			NossoNumero=cCarteira & "/" & cModalidade & NossoNumero & "-" & CalculaModulo11( cCarteira & cModalidade & cNossoNumero ,7 )
			Especial=false
			
			DebugResponse("Bradesco-NossoNumero: " & NossoNumero)

		CASE "341"	' Banco ITAÚ
			BancoDig="7"
			DebugResponse("ITAU: " & Carteira )
			
			AgenciaDig=""
			ContaDig=Mid( Conta, 7)
			Conta=Mid(Conta,1,5)
			
			cNossoNumero=Right( "00000000" & NossoNumero, 8 )
			cAgenciaNumero=Right( "0000" & Agencia, 4 )
			cContaNumero=Right( "00000" & Conta, 5 )
			cCarteira=Right( "000" & Carteira, 3 )
		
			if Carteira="126" or Carteira="131" or Carteira="146" or Carteira="150" or Carteira="168" then
		 	    cDAC=CalculaModulo10( cCarteira & cNossoNumero )
		    else
		        cDAC=CalculaModulo10( cAgenciaNumero & cContaNumero & cCarteira & cNossoNumero )
		    end if

			if Carteira="107" or Carteira="122" or Carteira="142" or Carteira="143" or Carteira="196" or Carteira="198" then
			
				cNumDoc=Right( "0000000" & NumeroDocumento, 7 )
				cCodCedente=Right("00000" & CodCedente,5)
				
				cLivre=cCarteira & _
				    cNossoNumero & _
					cNumDoc & _
					cCodCedente
				cLivre = cLivre  & CalculaModulo10( cLivre ) & "0"
				NumeroDocumento=cNumDoc & "-" &  CalculaModulo10(cNumDoc)
				
			else 
			    
				cLivre=cCarteira & _
					cNossoNumero & _
					cDAC & _
					cAgenciaNumero & _
					cContaNumero & _
					CalculaModulo10( cAgenciaNumero & cContaNumero ) & _
					"000"
					
			end if
			
			NossoNumero=cCarteira & "/" & NossoNumero & "-" & cDAC

		CASE "347"	'Sudameris
			BancoDig="6"
			'Página 10 e 16 da documentação - Cobrança sem registro
			
			AgenciaDig=""
			ContaDig=Mid(Conta,9)
			Conta=Mid(Conta,1,7)
			
			cNossoNumero=Right( "0000000000000" & NossoNumero, 13 )
			cAgenciaNumero=Right( "0000" & Agencia, 4 )
			cContaNumero=Right( "0000000" & Conta, 7 )
			
			cDAC=CalculaModulo10( cNossoNumero & cAgenciaNumero & cContaNumero )

			cLivre=cAgenciaNumero & _
				cContaNumero & _
				cDAC & _
				cNossoNumero
			'
			ContaDig=cDAC


		CASE "353", "033": ' Santander ou Banespa
            if Banco="353" then
                BancoDig="0"
            else
		        BancoDig="7"
		    end if

		    cCodCedente=Right("000000" & CodCedente, 7 )
		    cNossoNumero=Right("000000000000" & NossoNumero, 12)
            Especial=true
		    cDAC = CalculaModulo11(cNossoNumero,9)
            Especial=false
		
		    if Carteira="101" or Carteira="102" or Carteira="201" then
		        cLivre="9" & _
				    cCodCedente & _
				    cNossoNumero & cDAC & _
				    "0" & _
				    Carteira
		       'De acordo com a documentação fornecida o digito 0 fixo na posição 41
		       'chamado de IOF é referente apenas a seguradoras
		       NossoNumero=cNossoNumero & "-" &  cDAC
		    else
		        Response.Write("Carteira invalida!")
		        Response.End
		    end if
		    
		CASE "356"	'Real
			BancoDig="5"
			'Página 10 e 16 da documentação - Cobrança sem registro
			
			AgenciaDig=""
			ContaDig=Mid(Conta,9)
			Conta=Mid(Conta,1,7)

			cNossoNumero=Right( "0000000000000" & NossoNumero, 13 )
			cAgenciaNumero=Right( "0000" & Agencia, 4 )
			cContaNumero=Right( "0000000" & Conta, 7 )

			cDAC=CalculaModulo10( cNossoNumero & cAgenciaNumero & cContaNumero )

			cLivre=cAgenciaNumero & _
				cContaNumero & _
				cDAC & _
				cNossoNumero
			'
			ContaDig=cDAC
			
		CASE "389"	'Banco Mercantil
			BancoDig="1"
			
			cAgenciaNumero=Right( "0000" & Split(Agencia,"-")(0), 4 )
			cNossoNumero=Right( "00000000000" & NossoNumero, 11 )
			cCodCedente=Right( "000000000" & CodCedente, 9 )
			cModalidade=Right( "0" & Modalidade, 1 ) 'Indicador de desconto

			cLivre=cAgenciaNumero & _
				   cNossoNumero & _
				   cCodCedente & _
				   cModalidade
			
			'No codigo de barras o nosso numero te 9 posições, mas no calculo do digito do nosso numero há 10 posições
			cDAC=CalculaModulo11( cAgenciaNumero & "0" & cNossoNumero, 9 )
			NumeroDocumento=cNossoNumero & "-" &  cDAC

		CASE "399": ' HSBC
			BancoDig="9"
			
			' Verificar CNR ou CNR Facil

			AgenciaDig=""
			Agencia=Mid(Agencia,1,4)
			
			ContaDig=Mid(Conta,7,2)
			Conta=Mid(Conta,1,5)
			
			cCodCedente=Right("00000000" & CodCedente, 7 )
			
			'Utilizar o identificador 5 sempre que a data de vencimento estiver em branco e sem fator de vencimento. 
			if DataVencimento=DateValue( "2001/01/01" ) then
				Modalidade="5"
			end if
			
			DebugResponse( "Nosso Numero (Código do Documento): " & NossoNumero )
			DebugResponse( "Código do cedente: " & cCodCedente )
			DebugResponse( "Data de vencimento: " & DataVencimento )
			
			if Carteira="01" then	'Sem Registro
				cNossoNumero=Right( "0000000000000" & NossoNumero, 13 )
				cNossoNumero=cNossoNumero & CalculaModulo11(cNossoNumero,9)
				
				DebugResponse( "Nosso Numero (Com 1º Dígito): " & cNossoNumero )
				
				cTotal="0"
				cDataJuliana="0000"
				
				'Identificador 4: vincula vencimento, código do cedente e código do documento
				if Modalidade="4" then
				
					' Monta a Data Juliana do Venciento
					cDia=Right( "0" & Day( DataVencimento ), 2)
					cMes=Right( "0" & Month( DataVencimento ), 2)
					cAno=Right( Year( DataVencimento ), 2)
					cDataJuliana=cDia & cMes & cAno
					DebugResponse( "Data Juliana: " & cDataJuliana )
					
					' Efetua a soma Nosso numero (+Fim 4) + Cedente + Vencimento
					cNossoNumero=cNossoNumero & "4"
					cTotal=Soma( cDataJuliana, cCodCedente )
					cTotal=Soma( cNossoNumero, cTotal )
					cNossoNumero=cNossoNumero & CalculaModuloHSBC(cTotal)
					
					DebugResponse( "Total da Soma: " & cTotal )
					DebugResponse( "Nosso Numero (Com 2º Dígito): " & cNossoNumero )
				
					dStart=CDate("1/1/" & Year(DataVencimento) )
					nDias=DataVencimento-dStart+1
					cDataJuliana=Right( "000" & nDias, 3 ) & Right( Year(DataVencimento), 1 )
					
				else 'Identificador 5: vincula código do cedente e código do documento.
				
					' Efetua a soma Nosso Bumero (+Fim 5) + Cedente 
					cNossoNumero=cNossoNumero & "5"
					cTotal=Soma( cNossoNumero, cCodCedente )
					cNossoNumero=cNossoNumero & CalculaModuloHSBC(cTotal)
					
					DebugResponse( "Total da Soma: " & cTotal )
					DebugResponse( "Nosso Numero (Com 2º Dígito): " & cNossoNumero )

					cDataJuliana="0000"
				end if
				
				DebugResponse( "Data de Vencimento por Mês Juliano: " & cDataJuliana )
				cLivre=Right( cCodCedente, 7 ) & Right( "0000000000000" & NossoNumero, 13 ) & cDataJuliana & "2"
				cAgenciaContaExibicao=cCodCedente
				
			else 'Com Registro
			
				cAgenciaNumero=Right( "0000" & Agencia, 4 )
				cCodCedenteNumero=Right( "000000" & CodCedente, 7 )
				cCodConvenio=Right("00000" & Convenio, 5 )
				cNossoNumero=cCodConvenio & Right( "00000" & NossoNumero, 5 )
				cNossoNumero=cNossoNumero & CalculaModulo11( cNossoNumero, 7 )
				cLivre=cNossoNumero & cAgenciaNumero & cCodCedenteNumero & "001"
			
			end if
			
			' Atualiza o Nosso Número e a Carteira
			NossoNumero=cNossoNumero
			Carteira="CNR"
			
		CASE "409": ' UNIBANCO
			BancoDig="0"
			
			cCodCedente=Right("0000000" & CodCedente, 7 )
			if Modalidade="14" then
			    cNossoNumero=Right( "00000000000000" & NossoNumero, 14 )
			else
			    cNossoNumero=cCodCedente & Right( "0000000" & NossoNumero, 7 )
			end if
			
			cLivre="5" & _
				cCodCedente & _
				"00" & _
				cNossoNumero

            Especial=true
			cDV=CalculaModulo11(cNossoNumero,9)
			Especial=false
			
			cLivre=cLivre & cDV
			
			NossoNumero=cNossoNumero & "/" & cDV

		CASE "422": ' SAFRA
			BancoDig="7"
			
			cAgencia=Right( "000" & Split(Agencia,"-")(0), 5)
			cConta=Right( "00000000" & Replace(Conta,"-",""), 9)
			cNossoNumero=Right( "000000000" & NossoNumero, 8 )
			cDV=CalculaModulo11(cNossoNumero,9)
			
			cLivre="7" & _
				cAgencia & _
				cConta & _
				cNossoNumero & cDV
				
			cLivre = cLivre & CalculaModulo11(cLivre,9)
			NossoNumero=NossoNumero & "/" & cDV
		
		CASE "745": 'CITIBANK
		
            cNossoNumero = Right("00000000000" & NossoNumero, 11)
            cModalidade = Right("00" & Modalidade, 3)
            cCodCedente = Right("000000000" & CodCedente, 9)
            cDV = CalculaModulo11(cNossoNumero, 9)
            NossoNumero = cNossoNumero & "." & cDV
            cNossoNumero = cNossoNumero & cDV

            cLivre = "3" & cModalidade & cCodCedente & cNossoNumero
			
            
        CASE "756": ' SICOOB
			BancoDig="0"
			
			'Ver página 6 da documentação
			cCarteira=Right(Carteira, 1)                    'Código da carteira
			cConvenio=Right( "000" & Convenio, 4)           'Código da Cooperativa
			cModalidade=Right( "0" & Modalidade, 2)         'Modalidade
			cCodCedente=Right( "000000" & CodCedente, 7)    'Código do Cliente
			cNossoNumero=Right( "0000000" & NossoNumero, 7) 'Número do Título
			cParcela=Right( "000" & Request("Parcela"), 3)  'Número da Parcela
			
			cDV=CalculaDigCoob(cConvenio & "000" & cCodCedente & cNossoNumero)
			cNossoNumero=cNossoNumero & cDV
			
			cLivre=cCarteira & cConvenio & cModalidade & cCodCedente & cNossoNumero & cParcela
			
			cAgenciaContaExibicao=cConvenio & "/" & cCodCedente
			NossoNumero=NossoNumero & "-" & cDV

		CASE ELSE
		
			Response.Write( "<br>Banco não implementado<br><br>" )
			response.End()

	END SELECT
	
	DebugResponse("Campo Livre: <b>" & cLivre & "</b> Len: " & Len(cLivre) )

	cCodePadrao=cCodePadrao & cLivre
	DebugResponse("Campo Livre + CodPadrão: <b>" & cCodePadrao & "</b>" )
	
	nDig=CalculaModulo11( cCodePadrao, 9 )
	DebugResponse("DAC: <b>" & nDig & "</b>" )

	'Monta o digito verificador principal do boleto
	cCodePadrao=Mid( cCodePadrao, 1, 4 ) & nDig & Mid( cCodePadrao, 5 )
	
	DebugResponse("Código de Barra: <b>" & cCodePadrao & "</b> Len: " & Len(cCodePadrao) )
	CodBarras=cCodePadrao
	
	'Response.End()
	
        
END SUB

'== Faz parece ou não as mensagens de depuração
Sub DebugResponse( cStrValor ) 
	'Para exibir informações de depuração tire o comentário da linha abaixo
	Response.Write( cStrValor + "<br>" )
End Sub

'As rotinas de Data abaixo só são necessária para garantir compatibilidade com servidores Apache
'que não aceitam configurações LCID, caso contrario poderia-se usar a função padrão do ASP
'FormatDateTime( dData, 2) desde que tenha sido configurada para portugues
FUNCTION FormatData(dData)
	FormatData=Day(dData) & "/" & Month(dData) & "/" & Year(dData)
END FUNCTION

'A função abaixo substitui o que poderia ser escrito apenas como
'DateValue(Request(cDataVar ))
FUNCTION GetDataVar( cDataVar )
    GetDataVar = DateValue(Request(cDataVar ))
    
'   EMBORA AS LINHAS ABAXO GEREM COMPATIBILIDA DE COM LINUX, elas não funcionam no IIS7
'	Dim cValor
'	cValor=Request(cDataVar)
'	DebugResponse( cDataVar & ":" & cValor )
'	cValor=split(cValor,"/")
'	'Coloca no Formato Ano/Mes/Dia para criar a data
'	Dim cValor2
'	cValor2 = cValor(2) & "/" & cValor(1) & "/" & cValor(0)
'	GetDataVar = DateValue(CSTR(cValor2))
END FUNCTION

'== Soma numeros longos contidos em strings
FUNCTION Soma( A, B )
	Dim n,C,nC
	'Ajusta o tamanho das string
	if Len(A)>Len(B) then
		B=String(Len(A)-Len(B),"0") & B
	elseif Len(B)>Len(A) then
		A=String(Len(B)-Len(A),"0") & A
	end if
	
	nC=0 'Carry (vai um)
	C=""
		
	For n=Len(A) to 1 step -1
		n1=CInt(Mid(A,n,1))
		n2=CInt(Mid(B,n,1))
		n1=n1+n2+nC
		if n1>9 then
			nC=1
			n1=n1-10
		else
			nC=0
		end if
		C=n1 & C
	Next
	
	C=nC & C
	
	Soma=C
	
END FUNCTION

'== Calcula a Linha digitável ==
SUB CalculaLinhaDigitavel()
	Dim cCampo1, cCampo2, cCampo3, cCampo4, cCampo5
	
	cCampo1=Mid( CodBarras, 1, 4) & Mid( CodBarras, 20, 5)
	cCampo1=cCampo1 & CalculaModulo10(cCampo1)
	cCampo1=Mid( cCampo1, 1, 5) & "." & Mid( cCampo1, 6, 5)

	cCampo2=Mid( CodBarras, 25, 10)
	cCampo2=cCampo2 & CalculaModulo10(cCampo2)
	cCampo2=Mid( cCampo2, 1, 5) & "." & Mid( cCampo2, 6, 6)

	cCampo3=Mid( CodBarras, 35, 10)
	cCampo3=cCampo3 & CalculaModulo10(cCampo3)
	cCampo3=Mid( cCampo3, 1, 5) & "." & Mid( cCampo3, 6, 6)
            
	cCampo4 = Mid( CodBarras, 5, 1)

	cCampo5 = Mid( CodBarras, 6, 14)

	LinhaDigitavel=cCampo1 & " " & cCampo2 & " " & cCampo3 & " " & cCampo4 & " " & cCampo5

END SUB

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
		nTotal=nTotal+nNumero
		
		nMultiplicador=nMultiplicador+1
		if nMultiplicador>nBase then
			nMultiplicador=2
		end if
		
	Next
	
	if Len(cTexto)<43 and Banco="151" then 'Apenas para a nossa caixa
			'De acordo com a nota disponivel no manual do Banco Nossa Caixa na página 22
			'o resto não deve ser aproximado, e deveremos considerar o digito=0

		nResto=nTotal MOD 11

		if nResto=0 then
			CalculaModulo11=0
		elseif nResto>1 then
			CalculaModulo11=11-nResto
		else
			'Recalcular D2 com (D1+1)
			'Retorna um valor negativo informando que deve ser recalculado com incremento
			CalculaModulo11=-1
		end if

        DebugResponse("Modulo 11 (Nossa Caixa) Total: " & nTotal & " len: " & Len(cTexto) & " = " & CalculaModulo11)

	else ' rotinas padrão
		
		if Especial then
		
            'Logica Especial, para o Unibanco e alguns outros
		    nTotal=nTotal*10
		    nResultado=nTotal mod 11
    		
		    if nResultado=10 then '10 ou 0 (zero) pois "0" retorna "0" mesmo
		        if Banco="237" then
		            CalculaModulo11="P"
		        else
		            CalculaModulo11=0
		        end if
		    else
			    CalculaModulo11=nResultado
		    end if

            DebugResponse("Modulo 11 (Especial) Total: " & nTotal & " len: " & Len(cTexto) & " = " & CalculaModulo11)
		
		else 
		
		    'Padrão para todos os bancos
            'Logica antiga adotada por alguns bancos
		    'nResto=nTotal mod 11
		    'nResultado=11-nResto
		    'if nResultado=10 or nResultado=11 then
			'    CalculaModulo11=1
		    'else
			'    CalculaModulo11=nResultado
		    'end if

	        'Nova lógica (adotada pela maioria dos bancos)
	        nTotal=nTotal*10
	        nResultado=nTotal mod 11
	        if nResultado=0 or nResultado=10 then
		        CalculaModulo11=1
	        else
		        CalculaModulo11=nResultado
	        end if

            DebugResponse("Modulo 11 (Padrão) Total: " & nTotal & " len: " & Len(cTexto) & " = " & CalculaModulo11)

		end if
		
	end if
	
	DebugResponse("Modulo 11 Resultado: " & CalculaModulo11 )

END FUNCTION

'== Rotina para Calculo do digito do Nosso numero da Nossa Caixa==
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
	
	DebugResponse("CalculaPesos => Total: " & nTotal & " - Resto: " & nResto & " - DV: " & CalculaPesos )

END FUNCTION

FUNCTION CalculaDigCoob(cValor)
	DIM nContador,nV,nP,nD,nTotal,nResto
	DIM cPeso
	nTotal=0
	cPeso="319731973197319731973"
	dim cTotal
	cTotal="" 
	For nContador=1 to Len(cValor)
		
		nV=Int( Mid( cValor, nContador ,1) )
		nP=Int( Mid( cPeso, nContador ,1) )
		
		nD=(nV*nP)
		
		cTotal=cTotal & nD & " "
		nTotal=nTotal + nD
		
	Next
	
	nResto=nTotal MOD 11
	if nResto = 0 OR nResto = 1 then
		CalculaDigCoob=0
	else
		CalculaDigCoob=11-nResto
	end if
	
	DebugResponse("CalculaDigCoob => " & cTotal & " Total: " & nTotal & " - Resto: " & nResto & " - DV: " & CalculaDigCoob )

END FUNCTION


'== Rotina de calculo do Modulo Especial do HSBC ==
FUNCTION CalculaModuloHSBC( cTexto )
	DIM nContador, nNumero, nTotal
	DIM nMultiplicador, nResultado
	DIM cCaracter

	nTotal=0
	nMultiplicador=9
	
	For nContador=Len(cTexto) to 1 step -1
		
		cCaracter=Mid( cTexto, nContador ,1)
		nNumero=Int( cCaracter ) * nMultiplicador
		nTotal=nTotal+nNumero
		
		nMultiplicador=nMultiplicador-1
		if nMultiplicador<2 then
			nMultiplicador=9
		end if
		
	Next

	DebugResponse("Total:" & nTotal & " / 11")
	nResultado=nTotal MOD 11
		
	if nResultado=10 OR nResultado=0 then
		CalculaModuloHSBC=0
	else
		CalculaModuloHSBC=nResultado
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

if ContaDig<>"" then
	ContaDig="-" & ContaDig
end if

if cAgenciaContaExibicao="" then
    cAgenciaContaExibicao=Agencia & "/" & Conta & ContaDig
end if

%>
		<STRONG>
				Instruções:</STRONG><BR>
			<EM>1 - Imprima em impressora jato de tinta (ink jet) ou a laser em qualidade 
				normal ou alta. Não use modo económico.<BR>
				2 - Utilise folha A4 (210x297mm) ou carta (216x279mm) e margens mínimas a 
				esquerda e a direita do formulário<BR>
				3 - Corte na linha indicada. Não rasure, risque, fure ou dobre a região onde se 
				encontra o código de barras.</EM>
		<div align="center">
			<IMG src="<%=ImagePath%>corte.gif"><BR>
			<BR>
			<TABLE class="noborder" id="Table1" cellSpacing="0" cellPadding="0" width="640" border="0">
				<TR>
					<TD class="noborder" width="150">
						<img src="<%=ImagePath & Banco%>.jpg" border=0></TD>
					<TD class="noborder" vAlign="bottom">
						<TABLE class="noborder" id="Table8" height="25" cellSpacing="0" cellPadding="0" width="480"
							align="right" border="0">
							<TR>
								<TD class="noborder" style="BORDER-LEFT: #000000 2px solid" align="center" width="60">
									<font Class=campo style="FONT-SIZE: 10pt"><%=Banco & "-" & BancoDIG%></font></TD>
								<TD class="noborder" style="BORDER-LEFT: #000000 2px solid" align="right">
									<font Class=campo><%=Linhadigitavel%></font></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE id="Table7" cellSpacing="0" cellPadding="0" width="640" border="0">
				<TR>
					<TD height=27 vAlign="top" width="340" colSpan="3">Cedente<BR>
						<font Class=campo><%=Cedente%></font></TD>
					<TD vAlign="top" width="150">Agência/Código Cedente<BR>
						<DIV align="right" Class=campo><%=cAgenciaContaExibicao%></DIV>
					</TD>
					<TD bgcolor="#dddddd" vAlign="top" width="150">Vencimento<BR>
						<DIV align="right" Class=campo>
							<%
							if DataVencimento=DateValue( "2001/01/01" ) then
								Response.Write("contra apresentação")
							else
								Response.Write(FormatData(DataVencimento))
							end if
							%></DIV>
					</TD>
				</TR>
				<TR>
					<TD height=27 vAlign="top" colSpan="3">Sacado<BR>
						<font Class=campo><%=Sacado%></font></TD>
					<TD vAlign="top">Número do Documento<BR>
						<DIV align="right" Class=campo>
							<%=NumeroDocumento%></DIV>
					</TD>
					<TD vAlign="top" nowrap>Nosso Número<BR>
						<DIV align="right" Class=campo>
							<%=NossoNumero%></DIV>
					</TD>
				</TR>
				<TR>
					<TD height=27 vAlign="top" bgcolor="#dddddd" width="90">Espécie<BR>
						<DIV align="center" class=campo>R$</DIV>
					</TD>
					<TD vAlign="top" width="100">Quantidade
						<BR>
					</TD>
					<TD vAlign="top" width="150">(x) Valor<BR>
					</TD>
					<TD vAlign="top">(-) Descontos / Abatimentos<BR>
						&nbsp;</TD>
					<TD vAlign="top" bgcolor="#dddddd">(=) Valor Documento<BR>
						<DIV align="right" Class=campo>
							<%=FormatNumber( Valor, 2, true )%></DIV>
					</TD>

				</TR>
				<TR>
					<TD height=27 vAlign="bottom" colSpan="3">Demonstrativo:</TD>
					<TD vAlign="top">(+) Outros Acréscimos<BR>
						&nbsp;</TD>
					<TD vAlign="top">(=) Valor Cobrado<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD height=70 vAlign="top" colSpan="5"><font class=campo><%=Demonstrativo%></font></TD>
				</TR>
			</TABLE>
			<BR>
			<IMG src="<%=ImagePath%>corte.gif"><BR>
			<BR>
			<TABLE class="noborder" id="Table10" cellSpacing="0" cellPadding="0" width="640" border="0">
				<TR>
					<TD class="noborder" width="150">
						<img src="<%=ImagePath+Banco%>.jpg" border=0></TD>
					<TD class="noborder" vAlign="bottom">
						<TABLE class="noborder" id="Table11" height="25" cellSpacing="0" cellPadding="0" width="480"
							align="right" border="0">
							<TR>
								<TD class="noborder" style="BORDER-LEFT: #000000 2px solid" align="center" width="60">
									<font Class=campo style="FONT-SIZE: 10pt"><%=Banco%>-<%=BancoDig%></font></TD>
								<TD class="noborder" style="BORDER-LEFT: #000000 2px solid" align="right">
									<font Class=campo><%=Linhadigitavel%></font></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE id="Table9" cellSpacing="0" cellPadding="0" width="640" border="0">
				<TR>
					<TD height=27 colSpan="6">Local de Pagamento<BR>
						<font class=campo>Até o vencimento pagável em bancos do sistema de compensação</font></TD>
					<TD bgcolor="#dddddd" width="150">Vencimento<BR>
						<DIV align="right" Class=campo>
							<%
							if DataVencimento=DateValue( "2001/01/01" ) then
								Response.Write("contra apresentação")
							else
								Response.Write(FormatData(DataVencimento))
							end if
							%></DIV>
					</TD>
				</TR>
				<TR>
					<TD valign=top height=27 colSpan="6">Cedente<BR>
					<font Class=campo>
						<%=Cedente%></font></TD>
					<TD valign=top>Agência/Código Cedente<BR>
						<DIV align="right" Class=campo>
							<%=cAgenciaContaExibicao%></DIV>
					</TD>
				</TR>
				<TR>
					<TD valign=top height=27>Data Documento<BR>
						<DIV align="center" Class=campo>
							<%=FormatData( DataDocumento )%></DIV>
					</TD>
					<TD valign=top colSpan="2">Número do Documento<BR>
						<DIV align="center" Class=campo>
							<%=NumeroDocumento%></DIV>
					</TD>
					<TD valign=top >Espécie Doc.<BR>
						<DIV align="center" class=campo>RC</DIV>
					</TD>
					<TD valign=top>Aceite<BR>
						<DIV align="center" class=campo>N</DIV>
					</TD>
					<TD valign=top>Data Processamento<BR>
						<DIV align="center" Class=campo>
							<%=FormatData( Now )%></DIV>
					</TD>
					<TD valign=top nowrap>Nosso Número<BR>
						<DIV align="right" Class=campo>
							<%=NossoNumero%></DIV>
					</TD>
				</TR>
				<TR>
					<TD height=27 valign=top >Uso do Banco<BR>
						</TD>
					<TD valign=top>Carteira<BR>
						<DIV align="center" class=campo>
							<%=Carteira%></DIV>
					</TD>
					<TD valign=top bgcolor="#dddddd">Espécie<BR>
						<DIV align="center" class=campo>
							R$</DIV>
					</TD>
					<TD valign=top colSpan="2">Quantidade<BR></TD>
					<TD valign=top>(x) Valor<BR></TD>
					<TD valign=top bgcolor="#dddddd">(=) Valor documento<BR>
						<DIV align="right" Class=campo >
							<%=FormatNumber( Valor, 2, true )%></DIV>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top" colSpan="6" rowSpan="5">
						<b>Instruções</b><BR>
						<font class=campo><%=Instrucoes%></font>
						</TD>
					<TD valign=top height=27>(-) Descontos / Abatimentos<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD valign=top height=27>(-) Outras Deduções<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD valign=top height=27>(+) Mora / Multa<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD valign=top height=27>(+) Outros Acrescimos<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD valign=top height=27>(=) Valor<BR>
						&nbsp;</TD>
				</TR>
				<TR>
					<TD height=70 valign=top colSpan="7">Sacado<BR>
					<font Class=campo>
						<%=Sacado & "-" & SacadoDOC%>
						<br><%=Endereco1%>
						<br><%=Endereco2%>
						<br><%=Endereco3%></font></TD>
				</TR>
			</TABLE>
			<TABLE class="noborder" id="Table12" cellSpacing="0" cellPadding="0" width="640" border="0">
				<TR>
					<TD class="noborder"><div align=right><small>Autenticação Mecânica/Ficha de Compensação</small></div>
						<%GeraCodBarras(CodBarras)%></TD>
				</TR>
			</TABLE>
			<BR>
			<IMG src="<%=ImagePath%>corte.gif">
		</div> 
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
   </body>
</html>