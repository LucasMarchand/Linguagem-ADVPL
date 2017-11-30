// #########################################################################################
// Modulo : Financeiro - Contas a Pagar
// Fonte  : SLCR970
// ---------+-------------------+-----------------------------------------------------------
// Data     | Autor             | Descricao
// ---------+-------------------+-----------------------------------------------------------
// 04/09/17 | Lucas Marchand    | Relatório de Títulos a Pagar por Boleto (Tela x Excel)
// ---------+-------------------+-----------------------------------------------------------
#include 'rwmake.ch'
#include 'topconn.ch'
#include 'protheus.ch'

User Function SLCR970()

Private cPerg := 'SLCR970'

Private aDados		:= { { "Do Vcto" 	    	, "", "", "mv_ch1", "D", 	08, 0, 0, "G", ""	, "mv_par01", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", ""	    , "" }, ;
						 { "Até o Vcto"     	, "", "", "mv_ch2", "D", 	08, 0, 0, "G", ""	, "mv_par02", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", ""	    , "" }, ;
 						 { "Do Fornec"	     	, "", "", "mv_ch3", "C", 	06, 0, 0, "G", ""	, "mv_par03", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "SA3"	    , "" }, ;
 						 { "Até o Fornec"     	, "", "", "mv_ch4", "C",    06, 0, 0, "G", ""	, "mv_par04", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "SA3"	    , "" }, ;
 						 { "Da Natureza"     	, "", "", "mv_ch5", "C", 	10, 0, 0, "G", ""	, "mv_par05", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "SED"	    , "" }, ;
 						 { "Até a Natureza"     , "", "", "mv_ch6", "C", 	10, 0, 0, "G", ""	, "mv_par06", ""		    , "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "SED"	    , "" }, ;
 						 { "Gerar em"     		, "", "", "mv_ch7", "N", 	 1, 0, 0, "C", ""	, "mv_par07", "Tela"		, "", "", "", "", "Excel"		, "", "", "", "", "Ambos"		, "", "", "", "", ""			, "", "", "", "", ""			, "", "", "", "SED"	    , "" } }

AjustaSx1( cPerg, aDados )
Private cString		:= 'SE2'
Private titulo		:= 'Relatório de Títulos a Pagar por Boleto'
Private cDesc1		:= OemToAnsi( 'Este programa emite um Relatório de títulos que devem ser pagos por boleto. ' )
Private cDesc2		:= OemToAnsi( '' )
Private cDesc3		:= OemToAnsi( '' )
Private tamanho		:= 'G' // Tamanho da largura em tela: P = 80 Caracteres 	| 	M = 132 Caracteres		|	G = 220 Caracteres		| 
Private aOrdem		:= {} 
Private aReturn		:= { 'Zebrado', 1, 'Administracao', 2, 2, 1, '', 1 }
Private nomeprog	:= 'SLCR970'
Private nLastKey	:= 0
Private m_pag		:= 1
Private cUserId		:= DtoS( Date() ) + Time() + cUserName

If !IsBlind()
	Pergunte( cPerg, .f. )
	Private wnrel	:= SetPrint( cString, nomeprog, cPerg, titulo, cDesc1, cDesc2, cDesc3, .f., aOrdem,, tamanho )
Else
	wnrel := SetPrint(cString,NomeProg,"",@titulo,cDesc1,cDesc2,cDesc3,.T.,aOrdem,.T.,Tamanho,,.T.,,"default.drv",.T.,.T.,"SLCR970.##r")  
EndIf

If nLastKey == 27
	Set Filter to
	Return
Endif

aReturn[5]	:= 1
SetDefault ( aReturn, cString,,,tamanho,2)

If nLastKey == 27
	Set Filter to
	Return
Endif 
                       
If mv_par07 == 1 .or. mv_par07 == 3		//Formato para tela
	//                   01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
	//                   0         1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17        18        19        20        21        22
	//                   99/99/9999   99       999       999999999   999    999999        XXXXXXXXXXXXXXXXXXXX   999,999,999.99   9999999     XXXXXX   XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX   XXXXXXXXXXXXXXXXXXXX
	Private cabec1   := 'Vcto Real    Filial   Prefixo   Nº Título   Tipo   Cód. Fornec   Fornecedor             Valor Título     Natureza    Pedido   Comprador                                  Alçada                     '
	//                   01234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890
	//                   0         1         2         3         4         5         6         7         8         9         10        11        12        13        14        15        16        17        18        19        20        21        22
	//
	Private cabec2   := ' '                 
EndIf                 

If mv_par07 == 2 .or. mv_par07 == 3		//Funções para gerar em excel                         
	//--------------------------------------------------------------+
	//   Declaração das variáveis para o FWMsExcel                  |
	//--------------------------------------------------------------+
	Private cDirTmp		:= GetTempPath()
	Private cWorkSheet	:= "Contas a pagar"	
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
	oExcel:AddColumn( cWorkSheet, titulo, "Vcto Real"	,	1, 4, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Filial"		,	1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Prefixo"		,  	1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Nº Título"	,   1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Tipo"		,   1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Cód. Fornec.",   1, 1, .F.)                    
	oExcel:AddColumn( cWorkSheet, titulo, "Fornecedor"	,   1, 1, .F.)                 
	oExcel:AddColumn( cWorkSheet, titulo, "Valor Título",   1, 3, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Natureza"	,   1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Pedido"		,   1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Comprador"	,   1, 1, .F.)
	oExcel:AddColumn( cWorkSheet, titulo, "Alçada"		,   1, 1, .F.)
EndIf

RptStatus( { |lEnd| RptDetail( @lEnd ) }, titulo )

Return

********************************************************************************
Static Function RptDetail( lEnd )

Local cCount1, cQuery, cOrder, cCount2  

cCount1	:= 	"	SELECT COUNT(*) nRecnos FROM (  "

cQuery	:=	"	SELECT  "
cQuery	+=	"		SUBSTRING(E2_VENCREA, 7, 2) + '/' + SUBSTRING(E2_VENCREA, 5, 2) + '/' + SUBSTRING(E2_VENCREA, 1, 4) AS 'VCTO_REAL',  "
cQuery	+=	"		E2_FILORIG AS 'FILIAL', E2_PREFIXO AS 'PREFIXO', E2_NUM AS 'NUM_TITULO', E2_TIPO AS 'TIPO',  "
cQuery	+=	"		E2_FORNECE AS 'COD_FORNEC', E2_NOMFOR AS 'FORNECEDOR', E2_VALOR AS 'VALOR', E2_NATUREZ AS 'NATUREZA',  "
cQuery	+=	"		ISNULL(C7_NUM, '-') AS 'PEDIDO', ISNULL(Y1_NOME, '-') AS 'COMPRADOR', ISNULL(AL_DESC, '-') AS 'ALCADA', C7_USER  "
cQuery	+=	"	FROM SE2010  "
cQuery	+=	"		LEFT JOIN SD1010  "
cQuery	+=	"			ON D1_FILIAL = E2_FILORIG  "
cQuery	+=	"			AND D1_FORNECE = E2_FORNECE  "
cQuery	+=	"			AND D1_LOJA = E2_LOJA  "
cQuery	+=	"			AND D1_DOC = E2_NUM  "
cQuery	+=	"			AND D1_SERIE = E2_PREFIXO  "
cQuery	+=	"			AND SD1010.D_E_L_E_T_ = ''  "
cQuery	+=	"		LEFT JOIN SC7010  "
cQuery	+=	"			ON D1_FILIAL = C7_FILIAL  "
cQuery	+=	"			AND D1_FORNECE = C7_FORNECE  "
cQuery	+=	"			AND D1_LOJA = C7_LOJA  "
cQuery	+=	"			AND D1_ITEMPC = C7_ITEM  "
cQuery	+=	"			AND D1_COD = C7_PRODUTO  "
cQuery	+=	"			AND D1_PEDIDO = C7_NUM  "
cQuery	+=	"			AND SC7010.D_E_L_E_T_ = ''  "
cQuery	+=	"		LEFT JOIN SY1010  "
cQuery	+=	"			ON C7_USER = Y1_USER  "
cQuery	+=	"			AND SY1010.D_E_L_E_T_ = ''  "
cQuery	+=	"		LEFT JOIN SAL010  "
cQuery	+=	"			ON AL_COD = C7_APROV  "
cQuery	+=	"			AND SAL010.D_E_L_E_T_ = ''  "
cQuery	+=	"		JOIN SA2010  "
cQuery	+=	"			ON A2_COD = E2_FORNECE  "
cQuery	+=	"			AND A2_LOJA = E2_LOJA  "
cQuery	+=	"			AND SA2010.D_E_L_E_T_ = ''  "
cQuery	+=	"	WHERE  "
cQuery	+=	"		SE2010.D_E_L_E_T_ = ''  "
cQuery	+=	"		AND E2_SALDO > 0  "
cQuery	+=	"		AND E2_TIPO IN ('NF', 'FT', 'INS', 'ISS', 'PA', 'TX')  "
cQuery	+=	"		AND E2_CODBAR = ''  "     // Sem código de barras não tem como pagar por boleto
cQuery	+=	"		AND A2_FORMPGT = '2'  "   // 2 = Boleto	 ;	1 = Depósito C/C
cQuery	+=	"		AND E2_VENCREA BETWEEN '"+ DTOS(mv_par01) +"' AND '"+ DTOS(mv_par02) +"'  "
cQuery	+=	"		AND E2_FORNECE BETWEEN '"+ mv_par03 +"' AND '"+ mv_par04 +"'  "
cQuery	+=	"		AND E2_NATUREZ BETWEEN '"+ mv_par05 +"' AND '"+ mv_par06 +"'  "
cQuery	+=	"	GROUP BY  "
cQuery	+=	"		E2_VENCREA, E2_FILORIG, E2_PREFIXO, E2_NUM, E2_TIPO, E2_FORNECE,  "
cQuery	+=	"		E2_NOMFOR, E2_VALOR, E2_NATUREZ, C7_NUM, Y1_NOME, AL_DESC, C7_USER  "
cOrder 	:=	"	ORDER BY  "
cOrder	+=	"		E2_VENCREA, E2_FILORIG, E2_NOMFOR  "
                                                                                          
cCount2 :=	" ) AS TMP	"
        

dbUseArea( .t., 	'TOPCONN', TCGenQry(,,cCount1 + cQuery + cCount2), 'TMP', .f., .t. )
SetRegua( TMP->nRecnos )
TMP->( dbCloseArea() )
dbUseArea( .t., 'TOPCONN', TCGenQry(,, cQuery + cOrder ), 'TMP', .t., .t. )
                                      

TMP->( dbGoTop() )  // se posiciona no ínicio da query..                                                           
While TMP->( !EOF() )  // enquanto não for o final da query, ele vai fazendo a impressão.	
	IncRegua()

	If lEnd
		@ pRow()+1, 001 pSay 'CANCELADO PELO OPERADOR'
		Exit
	EndIf 

    // Validações
	cPrefixo := TMP->PREFIXO	
	If EMPTY(cPrefixo)
		cPrefixo := '-'  
	EndIf	          
	                 
	cComprador := TMP->COMPRADOR
	cUser := TMP->C7_USER
	If !EMPTY(cUser)
		cComprador := UsrFullName(cUser) //Pega o nome completo do comprador pelo SIGACFG
	EndIf

	// Aqui eu começo a fazer a impressão dos parâmetros em ...    
	// TELA
	If mv_par07 == 1 .or. mv_par07 == 3		
		If pRow() > 65 .or. m_pag == 1
			Cabec( Titulo, cabec1, cabec2, nomeprog, tamanho )
		EndIf
		
		@ pRow()+1, 000 pSay TMP->VCTO_REAL
		@ pRow()  , 013 pSay TMP->FILIAL
		@ pRow()  , 022 pSay cPrefixo            // TMP->PREFIXO
		@ pRow()  , 032 pSay TMP->NUM_TITULO
		@ pRow()  , 044 pSay TMP->TIPO
		@ pRow()  , 051 pSay TMP->COD_FORNEC
		@ pRow()  , 065 pSay Capitalace(AllTrim(SubString(TMP->FORNECEDOR, 1, 20)))
		@ pRow()  , 088 pSay Transform(TMP->VALOR, "@E 999,999,999.00")
		@ pRow()  , 105 pSay AllTrim(SubString(TMP->NATUREZA, 1, 7))
		@ pRow()  , 117 pSay TMP->PEDIDO
		@ pRow()  , 126 pSay Capitalace(AllTrim(SubString(cComprador, 1, 40)))
		@ pRow()  , 169 pSay Capitalace(AllTrim(SubString(TMP->ALCADA, 1, 20)))
	EndIf
    
	// EXCEL
	If mv_par07 == 2 .or. mv_par07 == 3								 
		oExcel:AddRow( cWorkSheet ,titulo, {TMP->VCTO_REAL, TMP->FILIAL, cPrefixo, TMP->NUM_TITULO, TMP->TIPO, TMP->COD_FORNEC, Capitalace(TMP->FORNECEDOR), TMP->VALOR, TMP->NATUREZA, TMP->PEDIDO, Capitalace(cComprador), Capitalace(TMP->ALCADA)} )   
	EndIf
	
	TMP->( dbSkip() ) //pula linha
End

TMP->( dbCloseArea() )
    
// EXCEL          
If mv_par07 == 2 .or. mv_par07 == 3
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
	fErase(cArq)   //deleta arquivo do temp do usuario
	oExcelApp:Destroy() // Deleta o processo do excel para não ficar preso após fechar
EndIf	          

// TELA
If mv_par07 == 1 .or. mv_par07 == 3	
	Roda( 0, '', tamanho )	 
	Set Filter to	
	
	wnrel:lEscClose := .F.	
	
	If aReturn[ 5 ] == 1
		Set Printer to
		dbCommitAll()
		OurSpool( wnrel )
	EndIf      
EndIf

MS_FLUSH()

Return  