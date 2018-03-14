#include 'rwmake.ch' 
#include 'protheus.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCA730      || Autor: Lucas Rocha          || Data: 28/02/18   ||
||-------------------------------------------------------------------------||
|| Descrição: Liberação de fretes de venda e compra de milho	           ||		                                   
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/                                            

User Function SLCA730()

Local cPerg	 	:= 'SLCA730'

Private cQuery   	:= ''
Private cQuery1  	:= ''
Private aList	 	:= Array(0)   
Private oDlg
Private oList               
Private lOk  	 	:= .F.  
Private dVenctos 	:= cToD('')
Private oOk	   	 := Loadbitmap( GetResources(), 'LBOK' )
Private oNo	   	 := Loadbitmap( GetResources(), 'LBNO' )
Private oVerde   	:= Loadbitmap( GetResources(), 'BR_VERDE' )
Private oVermelho	:= Loadbitmap( GetResources(), 'BR_VERMELHO' )
Private aButPed  	:= Array(0)                                   

Private aDados    := { 	{ "Situação"			, "", "", "mv_ch1", "N",  01, 0, 0, "C", "" , "mv_par01", "Todos"   , "", "", "", "", "Pendentes"  , "", "", "", "", "Liberados"  	, "", "", "", "", ""      , "", "", "", "", ""      , "", "", "", ""      , "" }, ;
			{ "Data Digit. de"		, "", "", "mv_ch2", "D",  08, 0, 0, "G", "" , "mv_par02", ""        , "", "", "", "", ""      		, "", "", "", "", ""      		, "", "", "", "", ""      , "", "", "", "", ""      , "", "", "", ""      , "" }, ;
                       	{ "Data Digit. até"		, "", "", "mv_ch3", "D",  08, 0, 0, "G", "" , "mv_par03", ""        , "", "", "", "", ""      		, "", "", "", "", ""      		, "", "", "", "", ""      , "", "", "", "", ""      , "", "", "", ""      , "" }, ;
                       	{ "Origem"			, "", "", "mv_ch4", "N",  01, 0, 0, "C", "" , "mv_par04", "Todas"   , "", "", "", "", "Compra" 	 	, "", "", "", "", "Venda"		, "", "", "", "", ""      , "", "", "", "", ""      , "", "", "", ""      , "" }  }

aAdd( aButPed, { 'PMSINFO', {|| SelectAll( aList ) }, 'Marca/Desmarca Todos Pendentes' } )
aAdd( aButPed, { 'BUDGETY', {|| Libera() }, 'Libera e Altera Vencimentos' } )
aAdd( aButPed, { 'PMSINFO', {|| Legenda() }, 'Legenda' } )                                                           

u_SX1Ajusta( cPerg, aDados )

If !Pergunte( cPerg, .T. )
	Return
EndIf   

If mv_par04 == 1 .or. mv_par04 == 3
	cQuery += " SELECT E2_FILORIG FILIAL, E2_NUM NUMERO, E2_EMISSAO EMISSAO, E2_EMIS1 DIGIT, E2_VENCTO VENCIMENTO, E2_VENCREA VENC_REAL, E2_FORNECE TRANSP, E2_LOJA LOJA_TRANSP, A2_NOME NOME_TRANSP, "
	cQuery += "   E2_VALOR VALOR, E2_USUALIB, E2_DATALIB, 'VENDA' AS ORIGEM, E2_PREFIXO, E2_PARCELA, E2_TIPO "
	cQuery += " FROM SE2010 "
	cQuery += "   JOIN SD1010 "
	cQuery += "     ON E2_FILORIG = D1_FILIAL "
	cQuery += "     AND E2_NUM = D1_DOC "
	cQuery += "     AND E2_FORNECE = D1_FORNECE "
	cQuery += "     AND E2_LOJA = D1_LOJA "
	cQuery += "     AND SD1010.D_E_L_E_T_ = '' "
	cQuery += "   JOIN SD2010 "
	cQuery += "     ON D1_FILIAL = D2_FILIAL "
	cQuery += "     AND D1_NFSORI = D2_DOC "
	cQuery += "     AND D1_NFSORII = D2_ITEM "
	cQuery += "     AND D1_NFSORIS = D2_SERIE "
	cQuery += "     AND SD2010.D_E_L_E_T_ = '' "
	cQuery += "   JOIN SA2010 "
	cQuery += "     ON A2_COD = D1_FORNECE "
	cQuery += "     AND A2_LOJA = D1_LOJA "
	cQuery += "     AND SA2010.D_E_L_E_T_ = '' "
	cQuery += " WHERE SE2010.D_E_L_E_T_ = '' "
//	cQuery += "   AND E2_HIST LIKE '%FRETE - OPERACAO MILHO%' "
	cQuery += "   AND D2_TP = 'MR' "
	cQuery += "   AND E2_TIPO = 'FT' "
	cQuery += "   AND E2_EMIS1 BETWEEN '" + dToS(mv_par02) + "' AND '" + dToS(mv_par03) + "' " 
	
	If mv_par01 == 2
		cQuery += "   AND E2_USUALIB = '' "
		cQuery += "   AND E2_SALDO <> 0 "
	ElseIf mv_par01 == 3
		cQuery += "   AND E2_USUALIB <> '' " 
	EndIf
EndIf

If mv_par04 == 1
	cQuery += " UNION ALL "
EndIf

If mv_par04 == 1 .or. mv_par04 == 2
	cQuery += " SELECT F8_FILIAL FILIAL, E2_NUM NUMERO, E2_EMISSAO EMISSAO, E2_EMIS1 DIGIT, E2_VENCTO VENCIMENTO, E2_VENCREA VENC_REAL, F8_TRANSP TRANSP, F8_LOJTRAN LOJA_TRANSP, A2_NOME NOME_TRANSP, "
	cQuery += "   E2_VALOR VALOR, E2_USUALIB, E2_DATALIB, 'COMPRA' AS ORIGEM, E2_PREFIXO, E2_PARCELA, E2_TIPO "
	cQuery += " FROM SF8010 "
	cQuery += "   JOIN SE2010 "
	cQuery += "     ON F8_NFDIFRE = E2_NUM "
	cQuery += "     AND F8_FILIAL = E2_FILORIG "
	cQuery += "     AND F8_TRANSP = E2_FORNECE "
	cQuery += "     AND F8_LOJTRAN = E2_LOJA "
	cQuery += "     AND SE2010.D_E_L_E_T_ = '' "
	cQuery += "   JOIN SD1010 "
	cQuery += "     ON F8_FILIAL = D1_FILIAL "
	cQuery += "     AND F8_NFORIG = D1_DOC "
	cQuery += "     AND F8_SERORIG = D1_SERIE "
	cQuery += "     AND F8_FORNECE = D1_FORNECE "
	cQuery += "     AND F8_LOJA = D1_LOJA "
	cQuery += "     AND SD1010.D_E_L_E_T_ = ''   "
	cQuery += "   JOIN SA2010 "
	cQuery += "     ON F8_TRANSP = A2_COD "
	cQuery += "     AND F8_LOJTRAN = A2_LOJA "
	cQuery += "     AND SA2010.D_E_L_E_T_ = '' "
	cQuery += " WHERE SF8010.D_E_L_E_T_ = '' "
	cQuery += "   AND E2_EMIS1 BETWEEN '" + dToS(mv_par02) + "' AND '" + dToS(mv_par03) + "' " 
	cQuery += "	  AND D1_TP = 'MR' "     
	
	If mv_par01 == 2
		cQuery += "   AND E2_USUALIB = '' "
		cQuery += "   AND E2_SALDO <> 0 "
	ElseIf mv_par01 == 3
		cQuery += "   AND E2_USUALIB <> '' " 
	EndIf
EndIf
 
cQuery += " ORDER BY DIGIT DESC, VENC_REAL DESC, FILIAL, TRANSP, LOJA_TRANSP "

dbUseArea( .t., 'TOPCONN', TCGenQry(,, cQuery ), 'TMP', .t., .t. )

Count To nNum
If nNum == 0
	MsgInfo("Não foi encontrado nenhuma NF.")
	dbCloseArea("TMP")
	Return
End

DbGoTop()
While !EOF()     

	aAdd( aList, { 	.F. 				,;
			Iif( Empty(TMP->E2_USUALIB), .T., .F.) ,;
			TMP->FILIAL 		,;
			TMP->NUMERO  		,;
			TMP->EMISSAO  		,;
			TMP->DIGIT  		,;
			TMP->VENCIMENTO  	,;
			TMP->VENC_REAL  	,;
			TMP->TRANSP  		,;
			TMP->LOJA_TRANSP 	,;
			TMP->NOME_TRANSP  	,;
			TMP->VALOR  		,;
			TMP->E2_USUALIB  	,;
			TMP->E2_DATALIB  	,;
			TMP->ORIGEM 		,;
			TMP->E2_PREFIXO		,;
			TMP->E2_PARCELA		,;
			TMP->E2_TIPO		} )

	dbSkip()
End

dbCloseArea("TMP")

MontaTela() 		
		

Return
                    

//////////////////////////////////
// 	     Funções  		//
//////////////////////////////////
********************************************************************************
Static Function MontaTela()        

aSize	 := MsAdvSize()
aObjects := {}

AAdd( aObjects, { 100, 100, .t., .t. } )
AAdd( aObjects, { 100, 030, .t., .f. } )
aInfo := { aSize[ 1 ], aSize[ 2 ], aSize[ 3 ], aSize[ 4 ], 3, 3 }
aPosObj := MsObjSize( aInfo, aObjects )

Define msDialog oDlg Title 'LIBERAÇÃO DE FRETES DE MILHO' From aSize[7],0 to aSize[6],aSize[5] OF oMainWnd PIXEL

@ aPosObj[1,1]+000,aPosObj[1,2]  ListBox oList ; //45,05
Fields Header  ' ', ' ', 'Filial', 'Título', 'Dt. Emissão', 'Dt. Digitação', 'Vencimento', 'Venc. Real', 'Cód. Transp.', 'Loja', 'Nome', 'Valor', 'Usuário Lib.', 'Data Lib.', 'Origem' ;
Size aPosObj[1,4]-aPosObj[1,2],aPosObj[1,3]-aPosObj[1,1] OF oDlg PIXEL ColSizes 60,30 ; //555,148
Pixel Of oDlg ;
On dblClick( aList:=SelectBox( oList:nAt, aList ), oList:Refresh() )

oList:aColSizes := { 10 , 10 , 25 , 40 , 35 , 35 , 35 , 35 , 40 , 20 , 90 , 35 , 50 , 35 , 40 }
				
oList:SetArray( aList )

oList:bHeaderClick := { |oObj,nCol| If( nCol==1, SelectAll( aList ), Nil), oList:Refresh() }

oList:bLine := { || { ;	
			Iif( aList[ oList:nAt, 01 ], oOk, oNo ), ;
			Iif( aList[ oList:nAt, 02 ], oVerde, oVermelho ),;                  
			aList[ oList:nAt, 03 ],  ;
			aList[ oList:nAt, 04 ],	;
			sToD( aList[ oList:nAt, 05 ] ),  	;
			sToD( aList[ oList:nAt, 06 ] ),  	;
			sToD( aList[ oList:nAt, 07 ] ),  	;
			sToD( aList[ oList:nAt, 08 ] ),  	;
			aList[ oList:nAt, 09 ],  ;
			aList[ oList:nAt, 10 ],  ;
			aList[ oList:nAt, 11 ],  ;
			Transform(aList[ oList:nAt, 12 ], "@E 999,999.99"), ;
			aList[ oList:nAt, 13 ],  ;
			sToD( aList[ oList:nAt, 14 ] ),  	;
			aList[ oList:nAt, 15 ]   ;      					
}} 

oList:nAt:=1
oList:Refresh()

Activate msDialog oDlg Centered On Init EnchoiceBar( oDlg, { || If( Libera(), lOk := .T., lOk := .F. ), If( lOk, oDlg:End(), Nil ) }, { || aList := Array( 0 ), oDlg:End() },, aButPed)

Return lOk    

********************************************************************************
Static Function SelectBox( nIt, aVector )

If !aVector[ nIt, 1 ] .AND. Empty(aVector[ nIt, 13 ])
	aVector[ nIt, 1 ] := .t.
	
Else	                 
	aVector[ nIt, 1 ] := .f.

EndIf
    
oList:Refresh()

Return( aVector ) 

********************************************************************************
Static Function SelectAll( aVector )

For i := 1 To Len(aVector)
	If !aVector[ i, 1 ] .AND. Empty(aVector[ i, 13 ])
		aVector[ i, 1 ] := .t.
	Else
		aVector[ i, 1 ] := .f.   	
	EndIf
Next i

oList:Refresh()

Return( aVector )       

********************************************************************************
Static Function Confirma()

Local lControle  :=  .F.

For i := 1 To Len(aList)
	If aList[i][1]  ==  .T.
		lControle  :=  .T.
        Exit
	EndIf
Next i

If !lControle
	MsgInfo("Você não selecionou nenhum título.")
	Return .F.	
EndIf


Return .T.

********************************************************************************
Static Function Libera()    

Local dProxima

If !Confirma() .OR. !AlteraVenctos()
   Return .F.
EndIf

For i := 1 To Len( aList )
  
	If aList[i][1]  // Se estiver marcado
		
		dbSelectArea("SE2")
		SE2->( dbSetOrder(1) ) 			//  E2_PREFIXO       E2_NUM          E2_PARCELA       E2_TIPO          E2_FORNECE      E2_LOJA  

		If SE2->( dbSeek( xFilial( "SE2" ) + aList[ i, 16 ] + aList[ i, 4 ] + aList[ i, 17 ] + aList[ i, 18 ] + aList[ i, 9 ] + aList[ i, 10 ] , .F.) )  
		
			SE2->( RecLock( 'SE2', .f. ) )
							
			SE2->E2_DATALIB := dDataBase
			SE2->E2_USUALIB := cUserName
			SE2->E2_VENCTO	:= dVenctos
			SE2->E2_VENCREA := dVenctos  
						
			SE2->( MsUnlock() )
			
			// Procura se tem títulos da União	
			cQuery1 := "	select E2_EMISSAO, E2_FILORIG, E2_VENCTO, E2_VENCREA, E2_NUM, E2_NATUREZ, E2_PREFIXO, E2_PARCELA, E2_TIPO, E2_FORNECE, E2_LOJA,E2_CODRET  "
			cQuery1 += " from SE2010  "
			cQuery1 += " where E2_FILORIG = '" + SE2->E2_FILORIG + "'  "
			cQuery1 += " 	and E2_NUM = '" + SE2->E2_NUM + "'  "
			cQuery1 += " 	and E2_EMISSAO = '" + dToS(SE2->E2_EMISSAO) + "'  "
			cQuery1 += " 	and E2_TIPO = 'TX' "
			cQuery1 += " 	and SE2010.D_E_L_E_T_ = ' '  " 
			
			dbUseArea( .t., 'TOPCONN', TCGenQry(,,cQuery1), 'SE2TMP', .f., .t. )    
			
			while !SE2TMP->(Eof())   
			                             
				dProxima := MonthSum(dVenctos,1) 
				dProxima := cToD('20/'+ substr(dToS(dProxima) ,5,2)  + '/'+ substr(dToS(dProxima) ,3,2)    )
				dProxima := DataValida( dProxima, .F. )  // Pega o dia útil para trás 

				IF ( AllTrim( SE2TMP->E2_NATUREZ ) == '2080201' .and. ( AllTrim( SE2TMP->E2_CODRET ) == '0588' .or. AllTrim( SE2TMP->E2_CODRET ) == '3208' ) ) .or. ;
					( AllTrim( SE2TMP->E2_NATUREZ ) == '2080204' .or. AllTrim( SElucas2TMP->E2_NATUREZ ) == '2080205' .or. AllTrim( SE2TMP->E2_NATUREZ ) == '2080206' )
					
					SE2->( dbSetOrder( 1 ) )
					SE2->( dbSeek( xFilial( "SE2" ) + SE2TMP->E2_PREFIXO + SE2TMP->E2_NUM + SE2TMP->E2_PARCELA + SE2TMP->E2_TIPO + SE2TMP->E2_FORNECE + SE2TMP->E2_LOJA, .F. ) )
					
					RecLock( "SE2", .f. )
					
					SE2->E2_VENCTO 	:= dProxima
					SE2->E2_VENCREA := dProxima 
					
					MsUnlock()
				EndIf
				
				SE2TMP->(dbSkip())
			End                      
			
			SE2TMP->( dbCloseArea() )
			SE2->( DBCloseArea() ) 
			
	    	EndIf	  	    
	EndIf   
Next

dbUseArea( .t., 'TOPCONN', TCGenQry(,, cQuery ), 'TMP', .t., .t. )

aSize(aList, 0)

DbGoTop()
While !EOF()     

	aAdd( aList, { 	.F. 				,;
			Iif( Empty(TMP->E2_USUALIB), .T., .F.) ,;
			TMP->FILIAL 		,;
			TMP->NUMERO  		,;
			TMP->EMISSAO  		,;
			TMP->DIGIT  		,;
			TMP->VENCIMENTO  	,;
			TMP->VENC_REAL  	,;
			TMP->TRANSP  		,;
			TMP->LOJA_TRANSP 	,;
			TMP->NOME_TRANSP  	,;
			TMP->VALOR  		,;
			TMP->E2_USUALIB  	,;
			TMP->E2_DATALIB  	,;
			TMP->ORIGEM 		,;
			TMP->E2_PREFIXO		,;
			TMP->E2_PARCELA		,;
			TMP->E2_TIPO		} )
		
	dbSkip()
End

dbCloseArea("TMP")

oList:SetArray( aList )
oList:bLine := { || { ;	
			Iif( aList[ oList:nAt, 01 ], oOk, oNo ), ; 
			Iif( aList[ oList:nAt, 02 ], oVerde, oVermelho ), ;                  
			aList[ oList:nAt, 03 ], ;
			aList[ oList:nAt, 04 ],	;
			sToD( aList[ oList:nAt, 05 ] ), ;
			sToD( aList[ oList:nAt, 06 ] ), ;
			sToD( aList[ oList:nAt, 07 ] ), ;
			sToD( aList[ oList:nAt, 08 ] ), ;
			aList[ oList:nAt, 09 ],	;
			aList[ oList:nAt, 10 ], ;
			aList[ oList:nAt, 11 ], ;
			Transform(aList[ oList:nAt, 12 ], "@E 999,999.99"), ;
			aList[ oList:nAt, 13 ], ;
			sToD( aList[ oList:nAt, 14 ] ),	;
			aList[ oList:nAt, 15 ]	;
}} 

oList:Refresh()  
 
Return .T.

********************************************************************************
Static Function AlteraVenctos()  

Local oDlg1
Local nOpc	:= 0
 
@ 000, 000 to 160, 350 Dialog oDlg1 Title 'Informe o Vencimento'

@ 043,010 Say 'Vencimento'	Size 080,10  Of oDlg1 Pixel
@ 040,060 MsGet dVenctos 	Size 040,04	 Of oDlg1 Pixel
Activate Dialog oDlg1 Centered On Init EnchoiceBar( oDlg1, { || nOpc := 1, If( nOpc == 1, oDlg1:End(), oDlg1:End() ) }, { || nOpc := 0, oDlg1:End() },,,,,,,.F. )

If nOpc == 1 .AND. Empty(dVenctos)

	MsgInfo('Vencimento não Informado!')
	Return AlteraVenctos()
ElseIf nOpc == 0
	Return .F.
	
EndIf 

Return .T. 

********************************************************************************
Static Function Legenda()

Local aLegenda := {}

Aadd(aLegenda, {"BR_VERMELHO" ,'Título já liberado'})
Aadd(aLegenda, {"BR_VERDE"    ,'Título pendente de liberação'})

BrwLegenda('Status', '', aLegenda)

Return .T.
