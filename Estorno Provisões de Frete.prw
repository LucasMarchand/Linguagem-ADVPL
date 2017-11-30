#include 'protheus.ch'
#include 'rwmake.ch'
#INCLUDE "TOTVS.CH"
/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCA900      || Autor: Lucas Rocha           || Data: 21/11/2017||
||-------------------------------------------------------------------------||
|| Descrição: Estorno das Provisões de Frete                               ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function SLCA900()

Local aButPed  		:= Array(0)
Local cQuery		:= ''           
Local cPerg 		:= 'SLC900A'  //Perguntas SX1                     
Local lOk	 		:= .f.

Private aList		:= Array( 0 ) 
Private nTot        := 0
Private nNum		:= 0
Private oDlg1, oDlg2
Private nOpc		:= 0
Private oList
Private oOk	   		:= Loadbitmap( GetResources(), 'LBOK' )
Private oNo	   		:= Loadbitmap( GetResources(), 'LBNO' )
Private dDate  		:= cToD('')

aAdd( aButPed, { 'PMSINFO', {|| SelectAll( aList ) }, 'Marcar/Desmarcar Todos' } )     // Adicionado um botão no submenu 'Ações Relacionadas'
						
Define Font oFntGet1 Name 'Arial' Size 12,19 Bold  // Estilo de fonte na variável oFntGet1 
                       
ValidPerg(cPerg)	//Abre os parâmetros definidos pela função ValidPerg() antes de carregar a página

If !Pergunte( cPerg , .T. )		// Se não confirmar a tela  
	Return   	                // Termina
EndIf
  

// Reúne as provisões de frete que estão em aberto, parametrizadas pela data passada acima

cQuery := " SELECT D2_FILIAL, D2_DOC AS DOC, D2_SERIE AS SERIE, D2_EMISSAO AS DT, D2_CLIENTE, D2_LOJA, D2_COD, D2_ITEM AS ITEM, "
cQuery += " 	B1_DESC, D2_TOTAL, D2_PROVFRT AS D2_PROV_1, '' AS D2_PROV_2, '' AS D1_PROV_1 "
cQuery += " FROM SD2010, SF2010, SB1010 "
cQuery += " WHERE SF2010.D_E_L_E_T_ = '' "
cQuery += " 	AND F2_EMISSAO between '" + DTOS(mv_par01) + "' and '" + DTOS(mv_par02) + "' "
cQuery += " 	AND F2_FILIAL = D2_FILIAL "
cQuery += " 	AND F2_DOC = D2_DOC "
cQuery += " 	AND F2_SERIE = D2_SERIE "
cQuery += " 	AND F2_CLIENTE = D2_CLIENTE "
cQuery += " 	AND F2_LOJA = D2_LOJA "
cQuery += " 	AND SD2010.D_E_L_E_T_ = '' "
cQuery += " 	AND B1_COD = D2_COD "
cQuery += " 	AND D2_ESTPFRT = '' "
cQuery += " 	AND D2_PROVFRT > 0 "

cQuery += " UNION ALL "

cQuery += " SELECT D2_FILIAL, D2_DOC AS DOC, D2_SERIE AS SERIE, D2_EMISSAO AS DT, D2_CLIENTE, D2_LOJA, D2_COD, D2_ITEM AS ITEM, "
cQuery += " 	B1_DESC, D2_TOTAL, '' AS D2_PROV_1, D2_PROVFR2 AS D2_PROV_2, '' AS D1_PROV_1 "
cQuery += " FROM SD2010, SF2010, SB1010 "
cQuery += " WHERE SF2010.D_E_L_E_T_ = '' "
cQuery += " 	AND F2_EMISSAO between '" + DTOS(mv_par01) + "' and '" + DTOS(mv_par02) + "' "
cQuery += " 	AND F2_FILIAL = D2_FILIAL "
cQuery += " 	AND F2_DOC = D2_DOC "
cQuery += " 	AND F2_SERIE = D2_SERIE "
cQuery += " 	AND F2_CLIENTE = D2_CLIENTE "
cQuery += " 	AND F2_LOJA = D2_LOJA "
cQuery += " 	AND SD2010.D_E_L_E_T_ = '' "
cQuery += " 	AND B1_COD = D2_COD "
cQuery += " 	AND D2_ESTPFR2 = '' "
cQuery += " 	AND D2_PROVFR2 > 0 "

cQuery += " UNION ALL "

cQuery += " SELECT D2_FILIAL, D1_DOC AS DOC, D1_SERIE AS SERIE, D1_DTDIGIT AS DT, D2_CLIENTE, D2_LOJA, D2_COD, D1_ITEM AS ITEM, "
cQuery += " 	B1_DESC, D2_TOTAL, '' AS D2_PROV_1, '' AS D2_PROV_2, D1_PROVFRT AS D1_PROV_1 "
cQuery += " FROM SD2010, SF2010, SB1010, SD1010, SF4010 "
cQuery += " WHERE SF2010.D_E_L_E_T_ = '' "
cQuery += " 	AND F2_FILIAL = D2_FILIAL "
cQuery += " 	AND F2_DOC = D2_DOC "
cQuery += " 	AND F2_SERIE = D2_SERIE "
cQuery += " 	AND F2_CLIENTE = D2_CLIENTE "
cQuery += " 	AND F2_LOJA = D2_LOJA "
cQuery += " 	AND SD2010.D_E_L_E_T_ = '' "
cQuery += " 	AND B1_COD = D2_COD "
cQuery += " 	AND D2_TIPO = 'N' "
cQuery += " 	AND D2_FILIAL = D1_FILIAL "
cQuery += " 	AND D2_DOC = D1_NFORI "
cQuery += " 	AND D2_SERIE = D1_SERIORI "
cQuery += " 	AND D2_ITEM = D1_ITEMORI "
cQuery += " 	AND D2_CLIENTE = D1_FORNECE "
cQuery += " 	AND D2_LOJA = D1_LOJA "
cQuery += " 	AND SD1010.D_E_L_E_T_ = '' "
cQuery += " 	AND D1_DTDIGIT between '" + DTOS(mv_par01) + "' and '" + DTOS(mv_par02) + "' "
cQuery += " 	AND D1_ESTPFRT = '' "
cQuery += " 	AND D1_PROVFRT > 0 "
cQuery += " 	AND D1_FILIAL = F4_FILIAL "
cQuery += " 	AND D1_TES = F4_CODIGO "
cQuery += " 	AND SF4010.D_E_L_E_T_ = '' " 

cQuery += " ORDER BY "
cQuery += " 	D2_FILIAL, D2_DOC, D2_SERIE, DT, D2_CLIENTE, D2_LOJA, D2_COD, D2_ITEM, D2_TOTAL "

MsAguarde( { || dbUseArea( .t., 'TOPCONN', TCGenQry( , , cQuery ), 'TMP', .f., .t. )}, "", 'Procurando...' )

Count To nLin
If nLin == 0
	MsgInfo("Não encontrado provisões de frete")
	dbCloseArea("TMP")
	Return
End

DBGOTOP()
While !EOF()    
    aAdd( aList, {		.F. ,;
    			  TMP->D2_FILIAL, ;
    			  TMP->DOC, ; 
    			  TMP->SERIE, ;
    			  TMP->DT, ;
    			  TMP->D2_CLIENTE, ;
    			  TMP->D2_LOJA, ;
    			  TMP->D2_COD, ;
    			  TMP->B1_DESC, ;
    			  TMP->ITEM, ;
    			  TMP->D2_TOTAL, ;
    			  TMP->D2_PROV_1, ;
    			  TMP->D2_PROV_2, ;
    			  TMP->D1_PROV_1 })	 

	dbSelectArea('TMP')
	dbSkip()
End                

dbCloseArea("TMP") 

If Len(aList) > 0

	aSize	 := MsAdvSize()
	aObjects := {}
	
	AAdd( aObjects, { 100, 100, .t., .t. } )
	AAdd( aObjects, { 100, 030, .t., .f. } )
	aInfo := { aSize[ 1 ], aSize[ 2 ], aSize[ 3 ], aSize[ 4 ], 3, 3 }
	aPosObj := MsObjSize( aInfo, aObjects )
	
	Define msDialog oDlg1 Title 'Seleção de Fretes' From aSize[7],0 to aSize[6],aSize[5] OF oMainWnd PIXEL
	
	@ aPosObj[1,1]+000,aPosObj[1,2]  ListBox oList ; //45,05
	Fields Header  '     ' ,'Filial' , 'Documento', 'Série', 'Data', 'Cliente', 'Loja', 'Cód. Prod.', 'Descrição', 'Item', 'Total', 'SD2 - Prov 1', 'SD2 - Prov 2', 'SD1 - Prov 1' ;
	Size aPosObj[1,4]-aPosObj[1,2],aPosObj[1,3]-aPosObj[1,1] OF oDlg1 PIXEL ColSizes 30,30 ; //555,148
	Pixel Of oDlg1 ;
	On dblClick( aList:=SelectBox( oList:nAt, aList ), oList:Refresh(), oNum:Refresh(), oTot:Refresh() )
				
	@ aPosObj[2,1]+000,aPosObj[2,2] TO  aPosObj[2,3], aPosObj[2,4] LABEL "Outros Dados" OF oDlg1 PIXEL  
	
	@ aPosObj[2,1]+012,008 Say 'Quantidade de Notas: ' Font oFntGet1 Size 300,300 Pixel Of oDlg1
	@ aPosObj[2,1]+012,130 Say oNum Var nNum Font oFntGet1 Size 300,300 Pixel Colors CLR_HBLUE Of oDlg1  
	
	@ aPosObj[2,1]+012,308 Say 'Total: R$' Font oFntGet1 Size 300,300 Pixel Of oDlg1
	@ aPosObj[2,1]+012,350 Say oTot Var Transform(nTot, "@E 999,999,999.99") Font oFntGet1 Size 300,300 Pixel Colors CLR_HRED Of oDlg1
	
	
	oList:aColSizes := {5,		 15,	  60,		  15,	  40,	   40,		 15,	 30,		   140,		   15,	  35,     40,  	  		 40,			40  	    	}
					// ListBox | Filial | Documento | Serie	| Data	 | Cliente | Loja  | Cód Produto | Descrição | Item | Total | SD2 - Prov 1 | SD2 - Prov 2 | SD1 - Prov 1 
					
	oList:SetArray( aList )
	
	oList:bHeaderClick := { |oObj,nCol| If( nCol==1, SelectAll( aList ), Nil), oList:Refresh(), oNum:Refresh(), oTot:Refresh() }
	
	oList:bLine := { || { ;	
		Iif( aList[ oList:nAt, 01 ], oOk, oNo ), ; 				// ListBox
		aList[ oList:nAt, 02 ], ;  								// Filial                         
		aList[ oList:nAt, 03 ], ;  								// Documento
		aList[ oList:nAt, 04 ], ;								// Serie
		sTod(aList[ oList:nAt, 05 ]), ;	  						// Data
		aList[ oList:nAt, 06 ], ;								// Cliente
		aList[ oList:nAt, 07 ], ;								// Loja
		aList[ oList:nAt, 08 ], ;								// Código de Produto
		aList[ oList:nAt, 09 ], ;								// Descrição
		aList[ oList:nAt, 10 ], ;								// Item		
		Transform(aList[ oList:nAt, 11 ], "@E 999,999.99"), ;	// Total
		Transform(aList[ oList:nAt, 12 ], "@E 999,999.99"), ;   // SD2 Prov Frt
		Transform(aList[ oList:nAt, 13 ], "@E 999,999.99"), ;   // SD2 Prov Frt 2
		Transform(aList[ oList:nAt, 14 ], "@E 999,999.99"), ;   // SD1 Prov Frt
	}}
	
	oList:nAt:=1
	oList:Refresh()
	
	Activate msDialog oDlg1 Centered On Init EnchoiceBar( oDlg1, { || If ( Confirm( aList ), lOk := .t., lOk := .f. ) , If ( lOk == .t., oDlg1:End(), Nil ) }, { || aList := Array( 0 ), oDlg1:End() },, aButPed)   
	
	If lOk	               
	
        MSAguarde ({ || WriteData( aList ) }, "Aguarde", "Gravando dados no banco..." )  
                       		
	EndIf                                                                                                                                                                   		
Else                           

	MsgInfo("Provisões selecionadas") 
	
EndIf

GetDRefresh()

Return     

********************************************************************************
Static Function ValidPerg(cPerg)

Local aArea    := GetArea()
Local aAreaSX1 := SX1->(GetArea())

dbSelectArea("SX1")
dbSetOrder(1)
cPerg := Padr(cPerg,10)
aRegs:={}

// Grupo/Ordem/Pergunta              /p2/03/Var     /Tip/T/D/P/GSC/Va/V1        /D1/D2/D3/Ct/V2/D1/D2/D3/Ct/V3/D1/D2/D3/C3/V4/D1/D2/D3/Ct/V5/D1/D2/D3/Ct/F3/XG
aAdd (aRegs, {cPerg, "01", "Data de Emissão  De    ?", "","", 	"mv_ch1", "D", 6,0,0, 	"G","", "mv_par01",""    ,"","","","",""    ,"","","","",""      ,"","","","","","","","","","","","","","","","",""})
aAdd (aRegs, {cPerg, "02", "Data de Emissão  Até   ?", "","", 	"mv_ch2", "D", 6,0,0, 	"G","", "mv_par02",""    ,"","","","",""    ,"","","","",""      ,"","","","","","","","","","","","","","","","",""})

For i:=1 to Len(aRegs)
	If !dbSeek(cPerg+aRegs[i,2])
		RecLock("SX1",.T.)
		For j:=1 to FCount()
			If j <= Len(aRegs[i])
				FieldPut(j,aRegs[i,j])
			Endif
		Next
		MsUnlock()
	Endif
Next

RestArea(aAreaSX1)
RestArea(aArea)
Return                                      
               
**********************************************************************
Static Function SelectBox( nIt, aVector )

If !aVector[ nIt, 1 ]
	aVector[ nIt, 1 ] := .t.
	nNum++
	nTot += aVector[ nIt, 12 ]
	nTot += aVector[ nIt, 13 ] 
	nTot += aVector[ nIt, 14 ]		
Else	                 
	aVector[ nIt, 1 ] := .f.
	nNum--
	nTot -= aVector[ nIt, 12 ]
	nTot -= aVector[ nIt, 13 ]  	
	nTot -= aVector[ nIt, 14 ]
EndIf
    
oList:Refresh()
Return( aVector )   

**********************************************************************
Static Function SelectAll( aVector )

For i := 1 To Len(aVector)
	If !aVector[ i, 1 ]
		aVector[ i, 1 ] := .t.
		nNum++
		nTot += aVector[ i, 12 ]
		nTot += aVector[ i, 13 ]
		nTot += aVector[ i, 14 ]
	Else
		aVector[ i, 1 ] := .f.   
		nNum--
		nTot -= aVector[ i, 12 ]
		nTot -= aVector[ i, 13 ]
		nTot -= aVector[ i, 14 ]		
	EndIf
Next i

oList:Refresh()

Return( aVector )  
  
********************************************************************** 
Static Function Confirm ( aVector )

Local lProced  :=  .F.

For i := 1 To Len(aVector)
	If aVector[i][1]  ==  .T.  // se está marcado
		lProced  :=  .T.
		Exit
	EndIf
Next

If lProced   
 
	Define MsDialog oDlg2 Title "Data de Estorno" From 074,060 To 200,248 Pixel Of oMainWnd
	
   //	@ 020,013 Say OemToAnsi("Data de estorno :") 		Size 020,018 of oDlg2 PIXEL
	@ 025,020 MsGet dDate           					Size 060,010 of oDlg2 PIXEL
	
	Activate Dialog oDlg2 Centered On Init EnchoiceBar( oDlg2, { || nOpc := 1 , oDlg2:End() }, { || nOpc := 2, oDlg2:End() },,,,,,,.F. )
	        
	If  nOpc == 1 .and. !Empty(dDate)
		Return .T.
	ElseIf  nOpc == 1 .and. Empty(dDate)
		MsgInfo("Você não selecionou a data de estorno.")
		Return Confirm( aVector )
	EndIF                
Else              
	MsgInfo("Você não selecionou nenhuma nota.")	
EndIf        

Return .F.  

********************************************************************** 
Static Function WriteData ( aList )

For i := 1 To Len( aList )
   
	If aList[i][1]	// Se estiver selecionado
							
    	If aList[i][14] > 0		// SD1 Prov Frt
    	
  			dbSelectArea( 'SD1' )
  			dbSetOrder( 1 )		//Filial + Doc + Serie + Fornecedor + Loja + Cod Produto + Item
  		 	dbSeek( aList[i][2] + aList[i][3] + aList[i][4] + aList[i][6] + aList[i][7] + aList[i][8] + aList[i][10] )
  		  	  		   			  		   		
 			If Found()
  		   		RecLock( 'SD1' , .F. )
				
				If aList[ i ][ 14 ] > 0		// D1_PROV_1  >  0
					SD1->D1_ESTPFRT  :=  'MANUAL - ' + cValToChar(dDate)
				EndIf	
					
  		   		MsUnlock()
 			Else
 				MsgAlert("Não foi possível se posicionar na SD1!") 
 				Return
 			EndIf
 			
 			dbCloseArea()
 			 		 			
 		Else    
 		
 			dbSelectArea( 'SD2' )
 			dbSetOrder( 3 )		//Filial + Doc + Serie + Cliente + Loja + Cod Produto + Item
  			dbSeek( aList[i][2] + aList[i][3] + aList[i][4] + aList[i][6] + aList[i][7] + aList[i][8] + aList[i][10] ) 
  			
  			If Found() 
  			 		   				   		
  		   		RecLock( 'SD2' , .F. )
  		   		
  		   		If aList[ i ][ 12 ] > 0 	// D2_PROV_1  >  0
   						SD2->D2_ESTPFRT  :=  'MANUAL - ' + cValToChar(dDate)
				ElseIf aList[ i ][ 13 ] > 0    // D2_PROV_2  >  0					
					SD2->D2_ESTPFR2  :=  'MANUAL - ' + cValToChar(dDate)						
				EndIf       
				
				MsUnlock() 
			Else
				MsgAlert("Não foi possível se posicionar na SD2!") 
				Return
 			EndIf 
 					
 			dbCloseArea()
 												
 		EndIf 
 	EndIf	
Next 

Return                                                                