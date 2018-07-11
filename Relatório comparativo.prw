#include 'rwmake.ch'
#include 'topconn.ch'
#include 'protheus.ch'
 
/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCR990      || Autor: Lucas Rocha           || Data: 24/05/18  ||
||-------------------------------------------------------------------------||
|| Descrição: Relatório comparativo das notas importadas da Target pela    ||
|| rotina SLCA990, com as notas lançadas no Protheus.					   ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||*/   

User Function SLCR990()

Private cPerg := 'SLCR990'

Private aDados		:= { { "Da Emissão"	    	, "", "", "mv_ch1", "D", 	08, 0, 0, "G", ""	, "mv_par01", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", ""	    , "" }, ;
						 { "Até Emissão"     	, "", "", "mv_ch2", "D", 	08, 0, 0, "G", ""	, "mv_par02", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", ""	    , "" } }
//						 { "Tipo de Relatório"	, "", "", "mv_ch3", "N", 	 1, 0, 0, "C", ""	, "mv_par03", "Resumido"	, "", "", "", "", "Comparativo"	, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", ""	    , "" } }

u_SX1Ajusta( cPerg, aDados )

Private cString		:= 'Z69'
Private titulo		:= 'Comparativo Target x Protheus'
Private cDesc1		:= OemToAnsi( '' )
Private cDesc2		:= OemToAnsi( '' )
Private cDesc3		:= OemToAnsi( '' )
Private tamanho		:= 'G' // P = 80 COLUNAS 	| 	M = 132 COLUNAS		|	G = 220 COLUNAS		| 
Private aOrdem		:= {} 
Private aReturn		:= { 'Zebrado', 1, 'Administracao', 2, 2, 1, '', 1 }
Private nomeprog	:= 'SLCR990'
Private nLastKey	:= 0       	

If !IsBlind()
	Pergunte( cPerg, .f. )
	Private wnrel	:= SetPrint( cString, nomeprog, cPerg, titulo, cDesc1, cDesc2, cDesc3, .f., aOrdem,, tamanho )
Else
	wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrdem,.T.,Tamanho,,.T.,,"default.drv",.T.,.T.,"SLCR990.##r")  
EndIf

If nLastKey == 27
	Set Filter to
	Return
Endif

aReturn[5]	:= 1
SetDefault( aReturn, cString,,,tamanho,2 )

If nLastKey == 27
	Set Filter to
	Return
Endif 
                                            
//--------------------------------------------------------------+
//   Declaração das variáveis para o FWMsExcel                  |
//--------------------------------------------------------------+
Private cDirTmp		:= GetTempPath()
Private cWorkSheet	:= "Target"	
Private oExcel	 	:= FWMSEXCEL():New() //instancia objeto Excel

//+---------------------------------------------------------------------------------------------------------------------------------------+
//|  Opções do objeto oExcel:AddColumn()                                                                                                  |
//|  oExcel():AddColumn(< cWorkSheet >, < titulo >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)-> NIL                               |
//|  Descrição: Adiciona uma coluna a tabela de uma WorkSheet                                                                             |
//+---------------------------------------------------------------------------------------------------------------------------------------+
//|  Nome		| Tipo		   | Descrição															| Default	 | Obrigatório |Referência|
//+---------------------------------------------------------------------------------------------------------------------------------------+
//|  cWorkSheet | Caracteres   | Nome da planilha													| 		 	 |     X	   |          | 
//|  titulo	    | Caracteres   | Nome da planilha													| 		 	 |     X	   |          | 
//|  cColumn	| Caracteres   | Titulo da tabela que será adicionada								| 		 	 |     X	   |          | 
//|  nAlign	    | Numérico     | Alinhamento da coluna ( 1-Left,2-Center,3-Right )					| 		 	 |     X	   |          | 
//|  nFormat	| Numérico     | Codigo de formatação ( 1-General,2-Number,3-Monetário,4-DateTime )	| 		 	 |     X	   |          | 
//|  lTotal	    | Lógico       | Indica se a coluna deve ser totalizada								| 		 	 |     X	   |          | 
//+---------------------------------------------------------------------------------------------------------------------------------------+

oExcel:AddWorkSheet	( cWorkSheet )		//Adiciona aba a tabela
oExcel:AddTable( cWorkSheet, titulo )	//Adiciona a tabela

oExcel:AddColumn( cWorkSheet, titulo, "Filial"				,	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Número NF"			,   1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Série"				,   1, 1, .F.)    
oExcel:AddColumn( cWorkSheet, titulo, "CNPJ do Fornecedor"	,	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Razão Social"		,   1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Chave"				,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Data Emissão"		,   1, 4, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "CFOP"				,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Total"				,   1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "ICMS"				,   1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "IPI"					,   1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "PIS"					,   1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Cofins"				,   1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Captura"				,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "Msg. XML"			,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Dt Digit"			,  	1, 4, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Canceladas"		,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Nº NF"				,  	1, 1, .F.)
//oExcel:AddColumn( cWorkSheet, titulo, "# Total"			,  	1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# ICMS"				,  	1, 3, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Terceiros"			,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Baixa/Transf."		,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Produtor"			,  	1, 1, .F.)
oExcel:AddColumn( cWorkSheet, titulo, "# Nota Prod."		,  	1, 1, .F.)                 

Processa( { |lEnd| RunProc( @lEnd ) }, titulo )

RptStatus( { |lEnd| RptDetail( @lEnd ) }, titulo )

Return

********************************************************************************
Static Function RunProc( lEnd )

Local cScript

ProcRegua( 10 )		// quantidade de IncProc no relatório

IncProc( 'Criando tabela no banco de dados...' )

// testa a existencia da tabela...
If !TCCanOpen( 'tSLCR990' )
	
	// cria a tabela...
	cScript := " CREATE TABLE tSLCR990 ( "
	cScript += "     T_FILIAL        VARCHAR(2), "
//	cScript += "     P_FILIAL        VARCHAR(2), "
	cScript += "     T_NUM           VARCHAR(9), "
//	cScript += "     P_NUM           VARCHAR(9), "
	cScript += "     T_SERIE         VARCHAR(3), "
//	cScript += "     P_SERIE         VARCHAR(3), "
	cScript += "     T_CGC           VARCHAR(14), "
	cScript += "     T_RAZAOSOCIAL   VARCHAR(40), "
	cScript += "     P_FORNECE       VARCHAR(6), "
	cScript += "     P_LOJA          VARCHAR(2), "
	cScript += "     T_CHAVE         VARCHAR(44), "
//	cScript += "     P_CHAVE         VARCHAR(44), "
	cScript += "     T_EMISSAO       VARCHAR(8), "
//	cScript += "     P_EMISSAO       VARCHAR(8), "
	cScript += "     T_CFOP          VARCHAR(5), "
//	cScript += "     P_CFOP          VARCHAR(5), "
	cScript += "     T_TOTAL         FLOAT(14), "
//	cScript += "     P_TOTAL         FLOAT(14), "
	cScript += "     T_ICMS          FLOAT(14), "
//	cScript += "     P_ICMS          FLOAT(14), "
	cScript += "     T_IPI           FLOAT(14), "
//	cScript += "     P_IPI           FLOAT(14), "
	cScript += "     T_PIS           FLOAT(14), "
//	cScript += "     P_PIS           FLOAT(14), "
	cScript += "     T_COFINS        FLOAT(14), "
//	cScript += "     P_COFINS        FLOAT(14), "
	cScript += "     T_CAPTURA       CHAR(1), "
	cScript += "     T_MSGXML        VARCHAR(40), "
	cScript += "     V_DTDIGIT		 VARCHAR(8), "
	cScript += "     V_SITUACAO		 CHAR(1), "
	cScript += "     V_NUMERO		 VARCHAR(9), "
//	cScript += "     V_TOTAL         FLOAT(14), "
	cScript += "     V_ICMS          FLOAT(14), "
	cScript += "     V_TERCEIROS	 CHAR(1), "
	cScript += "     V_PRODUTOR		 CHAR(1), "
	cScript += "     V_NOTAPROD		 VARCHAR(9) "
	cScript += " ) "

	SQLExec( cScript )
	
	cScript := "create index tSLCR9901 on tSLCR990 ( T_FILIAL, T_NUM, T_SERIE, T_CGC, T_EMISSAO ); create index tSLCR9902 on tSLCR990 ( T_FILIAL, T_CHAVE )"
	
	SQLExec( cScript )
	
EndIf

IncProc( 'Alimentando o relatório com os dados importados da Target' )
cScript := "  INSERT INTO tSLCR990 ( T_FILIAL, T_NUM, T_SERIE, T_CGC, P_FORNECE, P_LOJA, T_RAZAOSOCIAL, T_CHAVE, T_EMISSAO,  "
cScript += "      T_CFOP, T_TOTAL, T_ICMS, T_IPI, T_PIS, T_COFINS, T_CAPTURA, T_MSGXML, V_DTDIGIT, V_SITUACAO, V_NUMERO,  "
cScript += "      V_ICMS, V_TERCEIROS, V_PRODUTOR, V_NOTAPROD )  "
cScript += "  SELECT Z69_FILIAL, Z69_NUM, Z69_SERIE, Z69_CGC, Z69_FORNE, Z69_LOJA, Z69_NOME, Z69_CHAVE, Z69_EMIS,  "
cScript += "      Z69_CFOP, Z69_TOTAL, Z69_VALICM, Z69_VALIPI, Z69_VALPIS, Z69_VALCOF, Z69_CAPTUR, Z69_MSGXML,  "
cScript += " 	  '', '', '', 0, '', '', ''  "
cScript += "  FROM Z69010  "
cScript += "  WHERE D_E_L_E_T_ = ''  "
cScript += "  	AND Z69_EMIS BETWEEN '" + dToS( mv_par01 ) + "' AND '" + dToS( mv_par02 ) + "' "
SQLExec( cScript )

IncProc( 'Comparando as chaves com o Protheus' )
cScript := " UPDATE tSLCR990 "
cScript += " SET V_DTDIGIT = F1_DTDIGIT, "
cScript += "     V_ICMS = F1_VALICM "
cScript += " FROM tSLCR990 "
cScript += "     JOIN SF1010 "
cScript += "         ON T_CHAVE = F1_CHVNFE "
cScript += "         AND T_CHAVE <> '' "
cScript += "         AND T_FILIAL = F1_FILIAL "
cScript += "         AND SF1010.D_E_L_E_T_ = '' "
cScript += "     JOIN SD1010  "
cScript += "         ON F1_FILIAL = D1_FILIAL  "
cScript += "         AND F1_DOC = D1_DOC  "
cScript += "         AND F1_FORNECE = D1_FORNECE  "
cScript += "         AND F1_LOJA = D1_LOJA  "
cScript += "         AND F1_SERIE = D1_SERIE  "
cScript += "         AND SD1010.D_E_L_E_T_ = ''  "
SQLExec( cScript )

IncProc( 'Verificando a situação das notas') // NFe = Autorizada / Cancelada / Renegada
cScript := " UPDATE tSLCR990	  "
cScript += " SET V_SITUACAO = C00_SITDOC  "
cScript += " FROM tSLCR990   "
cScript += "     JOIN C00010  "
cScript += "         ON T_FILIAL = C00_FILIAL   "
cScript += "         AND T_CHAVE = C00_CHVNFE  "
cScript += "         AND C00010.D_E_L_E_T_ = '' "
SQLExec( cScript )

IncProc( 'Comparando "Filial + Número + Forn. + Loja + Emissao"' )
cScript := " UPDATE tSLCR990	 "
cScript += " SET V_NUMERO = F1_DOC, "
cScript += "     V_ICMS = F1_VALICM "
cScript += " FROM tSLCR990  "
cScript += "     JOIN SF1010  "
cScript += "         ON T_NUM = F1_DOC "
cScript += "         AND T_FILIAL = F1_FILIAL "
cScript += "         AND P_FORNECE = F1_FORNECE "
cScript += "         AND P_LOJA = F1_LOJA "
cScript += "         AND T_EMISSAO = F1_EMISSAO "
cScript += "         AND SF1010.D_E_L_E_T_ = '' "
cScript += "     JOIN SD1010  "
cScript += "         ON F1_FILIAL = D1_FILIAL  "
cScript += "         AND F1_DOC = D1_DOC  "
cScript += "         AND F1_FORNECE = D1_FORNECE  "
cScript += "         AND F1_LOJA = D1_LOJA  "
cScript += "         AND F1_SERIE = D1_SERIE  "
cScript += "         AND SD1010.D_E_L_E_T_ = ''  "
SQLExec( cScript )

// Analisar utilidade deste
// IncProc( 'Relaciona o total' )
// cScript := " update tSLCR990	 "
// cScript += " set V_TOTAL = P.TOTAL "
// cScript += " from tSLCR990  "
// cScript += "     join (	 "
// cScript += "         SELECT Z69_FILIAL, Z69_NUM, Z69_SERIE, Z69_CGC, Z69_EMIS, Z69_TOTAL, SUM(D1_TOTAL) TOTAL "
// cScript += "         FROM Z69010 "
// cScript += "             JOIN SF1010 "
// cScript += "                 ON Z69_NUM = F1_DOC "
// cScript += "                 AND Z69_FILIAL = F1_FILIAL "
// cScript += "                 AND F1_DTDIGIT BETWEEN '" + dToS( MonthSub(mv_par01, 3) ) + "' AND '" + dToS( MonthSum(mv_par02, 3) ) + "' "
// cScript += "                 AND SF1010.D_E_L_E_T_ = '' "
// cScript += "             JOIN SD1010 "
// cScript += "                 ON F1_FILIAL = D1_FILIAL "
// cScript += "                 AND F1_DOC = D1_DOC "
// cScript += "                 AND F1_FORNECE = D1_FORNECE "
// cScript += "                 AND F1_LOJA = D1_LOJA "
// cScript += "                 AND F1_SERIE = D1_SERIE "
// cScript += "                 AND SD1010.D_E_L_E_T_ = '' "
// cScript += "         WHERE Z69010.D_E_L_E_T_ = '' "
// cScript += "         	AND Z69_EMIS BETWEEN '" + dToS( mv_par01 ) + "' AND '" + dToS( mv_par02 ) + "' "
// cScript += "         GROUP BY Z69_FILIAL, Z69_NUM, Z69_SERIE, Z69_CGC, Z69_EMIS, Z69_TOTAL "
// cScript += "         HAVING Z69_TOTAL = SUM(D1_TOTAL)  "
// cScript += "     ) P ON T_FILIAL = P.Z69_FILIAL "
// cScript += "         AND T_NUM = P.Z69_NUM "
// cScript += "         AND T_SERIE = P.Z69_SERIE "
// cScript += "         AND T_CGC = P.Z69_CGC "
// cScript += "         AND T_EMISSAO = P.Z69_EMIS "

IncProc('Verificando as notas de entrada de terceiros')
cScript := " UPDATE tSLCR990 "
cScript += " SET V_TERCEIROS = 'E' "
cScript += " WHERE T_CFOP < '5101' "
SQLExec( cScript )

IncProc('Verificando os produtores')
cScript := " UPDATE tSLCR990 "
cScript += " SET V_PRODUTOR = 'S' "
cScript += " FROM tSLCR990	 "
cScript += "     JOIN SA2010 "
cScript += " 		ON A2_COD = P_FORNECE "
cScript += " 		AND A2_LOJA = P_LOJA "
cScript += " 		AND SA2010.D_E_L_E_T_ = '' "
cScript += " 		AND A2_GRPTRIB = '002' "
SQLExec( cScript )

IncProc('Verificando as notas de produtor')
cScript := " UPDATE tSLCR990 "
cScript += " SET V_NOTAPROD = F1_DOC "
cScript += " FROM tSLCR990 "
cScript += "     JOIN SF1010 "
cScript += "         ON T_FILIAL = F1_FILIAL "
cScript += "         AND T_CHAVE = F1_CHVPROD "
cScript += "         AND SF1010.D_E_L_E_T_ = '' "
SQLExec( cScript )

// IncProc('')
// cScript := " update tSLCR990	 "
// cScript += " set V_NOTAPROD = B1_GRTRIB "
// cScript += " from tSLCR990  "
// cScript += "     join (	 "
// cScript += "         SELECT Z69_FILIAL, Z69_NUM, Z69_SERIE, Z69_CGC, Z69_EMIS, Z69_TOTAL, B1_GRTRIB "
// cScript += "         FROM Z69010 "
// cScript += "         	 JOIN SD1010 "
// cScript += "                 ON Z69_NUM = D1_NFP "
// cScript += "                 AND D1_DTDIGIT BETWEEN '" + dToS( MonthSub(mv_par01, 3) ) + "' AND '" + dToS( MonthSum(mv_par02, 3) ) + "' "
// cScript += "                 AND D1_FILIAL = Z69_FILIAL "
// cScript += "                 AND SD1010.D_E_L_E_T_ = '' "
// cScript += "             JOIN SF1010 "
// cScript += "                 ON F1_FILIAL = D1_FILIAL "
// cScript += "                 AND F1_DOC = D1_DOC "
// cScript += "                 AND F1_FORNECE = D1_FORNECE "
// cScript += "                 AND F1_LOJA = D1_LOJA "
// cScript += "                 AND F1_SERIE = D1_SERIE "
// cScript += "                 AND SF1010.D_E_L_E_T_ = '' "
// cScript += "             JOIN SB1010 "
// cScript += "                 ON B1_COD = D1_COD "
// cScript += "                 AND SB1010.D_E_L_E_T_ = '' "
// cScript += "         WHERE Z69010.D_E_L_E_T_ = '' "
// cScript += "         	AND Z69_EMIS BETWEEN '" + dToS( mv_par01 ) + "' AND '" + dToS( mv_par02 ) + "' "
// cScript += "         GROUP BY Z69_FILIAL, Z69_NUM, Z69_SERIE, Z69_CGC, Z69_EMIS, Z69_TOTAL, B1_GRTRIB  "
// cScript += "     ) P ON T_FILIAL = P.Z69_FILIAL "
// cScript += "         AND T_NUM = P.Z69_NUM "
// cScript += "         AND T_SERIE = P.Z69_SERIE "
// cScript += "         AND T_CGC = P.Z69_CGC "
// cScript += "         AND T_EMISSAO = P.Z69_EMIS "


Return

********************************************************************************
Static Function RptDetail( lEnd )

Local cCount1, cQuery, cOrder, cCount2, cScript
Local cValDigit, cValSituacao, cValNumero, /* cValTotal ,*/ cValIcms, cValTerceiros, cValTransf, cValProdutor, cValNotaProd

cCount1	:= 	" SELECT COUNT(*) nRecnos FROM ( "
cQuery	:=	"	SELECT * FROM tSLCR990 "
cOrder	:=	"	ORDER BY T_FILIAL, T_NUM, T_SERIE, T_EMISSAO "
cCount2	:= 	" ) AS TMP "

dbUseArea( .t., 'TOPCONN', TCGenQry(,, cCount1 + cQuery + cCount2 ), 'TMP', .f., .t. )
SetRegua( TMP->nRecnos )
TMP->( dbCloseArea() )

dbUseArea( .t., 'TOPCONN', TCGenQry(,, cQuery + cOrder ), 'TMP', .t., .t. )
TMP->( DBGoTop() )

While TMP->( !Eof() )
	If lEnd
		@ pRow()+1, 001 pSay 'CANCELADO PELO OPERADOR'
		Exit
	EndIf

	IncRegua()

	cValDigit 		:= IIF( Empty( TMP->V_DTDIGIT ), '#N/D', STOD( TMP->V_DTDIGIT ) )
	cValSituacao	:= IIF( Empty( TMP->V_SITUACAO ), '-', IIF( TMP->V_SITUACAO == '1', 'Autorizada', IIF( TMP->V_SITUACAO == '2', 'Denegada', IIF( TMP->V_SITUACAO == '3', 'Cancelada', 'Necessário nova parametrização' ) ) ) )
	cValNumero		:= IIF( Empty( TMP->V_NUMERO ), '#N/D', TMP->V_NUMERO )
//	cValTotal		:= IIF( Empty( TMP->V_TOTAL), '-', TMP->V_TOTAL )
	cValIcms		:= IIF( Empty(TMP->V_ICMS), '-', TMP->V_ICMS )
	cValTerceiros	:= IIF( TMP->V_TERCEIROS == 'E', 'Entrada', '#N/D' )
	cValTransf		:= IIF( 'SLC ALIMENTOS' $ TMP->T_RAZAOSOCIAL, 'Sim', '#N/D')
	cValProdutor	:= IIF( TMP->V_PRODUTOR == 'S', 'Sim', '#N/D' )
	cValNotaProd	:= IIF( Empty( TMP->V_NOTAPROD ), '#N/D', TMP->V_NOTAPROD )

	oExcel:AddRow( cWorkSheet, titulo, { ;
		TMP->T_FILIAL, TMP->T_NUM, TMP->T_SERIE, TMP->T_CGC, TMP->T_RAZAOSOCIAL, TMP->T_CHAVE, STOD(TMP->T_EMISSAO), ;
		TMP->T_CFOP, TMP->T_TOTAL, TMP->T_ICMS, TMP->T_IPI, TMP->T_PIS, TMP->T_COFINS, TMP->T_CAPTURA, TMP->T_MSGXML, ;
		cValDigit, cValSituacao, cValNumero, /* cValTotal ,*/ cValIcms, cValTerceiros, cValTransf, cValProdutor, cValNotaProd } )
	
	TMP->( dbSkip() )
End

TMP->( dbCloseArea() )

cScript := " DELETE tSLCR990; DROP TABLE tSLCR990 "
SQLExec( cScript )

If !ApOleClient('MsExcel')
	MsgAlert("Microsoft Excel não instalado!")
	Return
EndIf

oExcel:Activate()
cArq := CriaTrab( NIL, .F. ) + ".xls"
oExcel:GetXMLFile(cArq)
__CopyFile( cArq, cDirTmp + cArq )
oExcelApp := MsExcel():New()
oExcelApp:WorkBooks:Open( cDirTmp + cArq )
oExcelApp:SetVisible(.T.)
fErase(cArq)   //deleta arquivo da pasta temp do usuario
oExcelApp:Destroy() // Deleta o processo do excel para não ficar preso

MS_FLUSH()

Return

********************************************************************************
Static Function SQLExec( cScript )

If TCSQLExec( cScript ) < 0
	MsgInfo( TCSqlError(), 'Falha na execução do script no banco de dados.' )
EndIf

Return
