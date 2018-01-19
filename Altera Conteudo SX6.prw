#include 'rwmake.ch' 
#include 'protheus.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCF1050      || Autor: Lucas Rocha          || Data: 05/12/17  ||
||-------------------------------------------------------------------------||
|| Descrição: Ajuste de usuários nos parâmetros da SX6			   ||		                                   
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/                                            

User Function SLCF1060()

//////////////////////////////////
// 	   Variáveis  		//
//////////////////////////////////
Local cPerg	:= 'SLCF1060'
Local lCond 	:= ''
Local lCond1	:= ''
Local cString   := ''
Local cX6	:= ''
Local cDiv	:= ''
Local nOpc	:= 0
Local cDir    	:= '\logs\'
Local cArq    	:= 'LogSX6_' + DTOS( DATE() )  + '_' + SUBSTR(TIME(), 1, 2) + SUBSTR(TIME(), 4, 2) + SUBSTR(TIME(), 7, 2) + '.txt' 
Local nHandle 	:= 0
Local lAltera	:= .F.

Private mv1	 := 0
Private mv2_name := Space(30)
Private mv3_name := Space(30)
Private mv2_id	 := Space(6)
Private mv3_id	 := Space(6)
Private cDiv1  	 := '' 
Private cLog	 := '' 
Private aList	 := Array(0)   
Private oDlg1, oDlg2
Private nOpc   	 := 0
Private oList
Private oOk	 := Loadbitmap( GetResources(), 'LBOK' )
Private oNo	 := Loadbitmap( GetResources(), 'LBNO' )
Private aButPed  := Array(0) 
Private lGok

//////////////////////////////////
// 	 Desenvolvimento 	//
//////////////////////////////////
ValidPerg( cPerg )
If !Pergunte( cPerg, .T. )
	Return
EndIf 
       
mv1 := mv_par01		// Função
// (mv_par02)		// Usuário Com Parâmetros
// (mv_par03)		// Usuário Sem Parâmetros

If !F1060ValidParam() .OR. !F1060PegaUsr()
	Return U_SLCF1060()
EndIf 

// Se os parâmetros estão Ok começa a busca na SX6      
dbSelectArea('SX6')               
SX6->( dbSetOrder(1) )

// Pesquisa pelo NOME do usuário
lCond	:= ' "' + mv2_name + '" $ X6_CONTEUD '
    
SX6->( dbSetFilter( { || &lCond }, lCond) )

SX6->( dbGoTop() )

// Se encontrou algum varre 
WHILE !EOF()
	lAltera	:= .T.
	cX6 	:= AllTrim( X6_CONTEUD )
	cDiv1 	:= ''
	cDiv 	:= F1060PegaDiv( cX6 )	// Pega o caractere que separa os usuários    
         
	If ( mv1 == 1 ) .AND. !( mv3_name $ cX6 )	// Substituir
	   
		cString := StrTran( cX6, mv2_name, mv3_name )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_name $ cX6 )	// Espelhar
	
		If SUBSTR( cX6, AT(mv2_name, cX6) + LEN(mv2_name), LEN('@slcalimentos') ) == '@slcalimentos'   // Se estiver gravado como e-mail	    	
	    	cString := cX6 + cDiv1 + mv3_name + '@slcalimentos.com.br' + cDiv  
	    	
	 	Else		
			cString := cX6 + cDiv1 + mv3_name + cDiv 
						       
		EndIf
											
	ElseIf ( mv1 == 3 )	// Excluir    

		If SUBSTR( cX6, AT(mv2_name, cX6) + LEN(mv2_name) , LEN('@slcalimentos') ) == '@slcalimentos' 						
			cString := StrTran( cX6, mv2_name + '@slcalimentos.com.br' + cDiv, '' ) 	// Se possuir divisor irá excluir 
			
			If cX6 == cString   // Se não possui divisor a StrTran não funcionou
				cString := StrTran( cX6, mv2_name + '@slcalimentos.com.br', '' ) 
				
			EndIf
		Else	
			cString := StrTran( cX6, mv2_name + cDiv, '' )  
			
			If cX6 == cString
				cString := StrTran( cX6, mv2_name, '' )
				
			EndIf
		EndIf		
	Else
		lAltera := .F.
	
	EndIf
			
	If lAltera

		aAdd( aList, { .F. , X6_FIL, X6_VAR, X6_TIPO, X6_DESCRIC, X6_DESC1, X6_DESC2, X6_CONTEUD, cString } ) 		
	EndIf
	 	 
	dbSkip()
END

SX6->( dbClearFilter() )
 

// Pesquisa pelo ID do usuário
lCond	:= ' "' + mv2_id + '" $ X6_CONTEUD '
  
SX6->( dbSetFilter( { || &lCond }, lCond) )

SX6->( dbGoTop() )

// Varre todos parâmetros achados 
WHILE !EOF()   
	lAltera	:= .T.
	cX6 	:= AllTrim( X6_CONTEUD )
	cDiv1 	:= ''	
	cDiv 	:= F1060PegaDiv( cX6 )		// Pega o caractere que separa os usuários
    
         
	If ( mv1 == 1 ) .AND. !( mv3_id $ cX6 )		// Substituir
		cString := StrTran( cX6, mv2_id, mv3_id )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_id $ cX6 ) 	// Espelhar
		cString := cX6 + cDiv1 + mv3_id + cDiv			       
											
	ElseIf ( mv1 == 3 )				// Excluir
		cString := StrTran( cX6, mv2_id + cDiv, '' )
		
		If cX6 == cString
			cString := StrTran( cX6, mv2_id, '' )
			
		EndIf
	Else
		lAltera	:= .F.
		
	EndIf		
	
	If lAltera
	
		aAdd( aList, { .F. , X6_FIL, X6_VAR, X6_TIPO, X6_DESCRIC, X6_DESC1, X6_DESC2, X6_CONTEUD, cString } ) 	 		
	EndIf
 
	dbSkip()
END                     

SX6->( dbClearFilter() )
SX6->( dbCloseArea() )


// Criar tela
If Len(aList) > 0
	If F1060MontaTela() 		
		MSAguarde ({ || F1060Grava() }, "Aguarde", "Gravando dados no banco..." )
		If lGok	// Se foi gravado corretamente				
			// Salvar cLog ==> Na pasta '\TOTVS 12\Microsiga\Protheus_Data\logs\LogSX6_Data_Hora'
			nHandle := FCreate( cDir + cArq )
			
			If nHandle < 0
				MsgAlert( 'Erro durante a gravação do log. Verifique se existe o caminho: ' + cDir + cArq  )
				
			Else
				FWrite( nHandle, cLog )
				FClose( nHandle )
				MsgInfo( 'Log salvo no diretório "Protheus_Data' + cDir + cArq + '"' ) 
				
			EndIf   		
		EndIf
	Else
		Return U_SLCF1060()
		
	EndIf
	
Else
	MsgInfo( 'Não existem parâmetros para serem configurados.' )   

	Return U_SLCF1060()			   
EndIf
 

Return  // FIM
                    

//////////////////////////////////
// 	     Funções  		//
//////////////////////////////////
********************************************************************************
Static Function ValidPerg(cPerg)  	// Cria o grupo de perguntas 

Local aArea    := GetArea()                                                                              
Local aAreaSX1 := SX1->(GetArea())

dbSelectArea("SX1")
dbSetOrder(1)
cPerg := Padr(cPerg,10)
aRegs:={}

// Grupo/Ordem/Pergunta              /p2/03/Var     /Tip/T/D/P/GSC/Va/V1        /D1/D2/D3/Ct/V2/D1/D2/D3/Ct/V3/D1/D2/D3/C3/V4/D1/D2/D3/Ct/V5/D1/D2/D3/Ct/F3/XG                                            
aAdd (aRegs, {cPerg, "01", "O que fazer	:				", "","", 	"mv_ch1", "N",  1,0,0, 	"C","","mv_par01","Substituir"	,"","","","","Espelhar"		,"","","","","Excluir"	,"","","","","","","","","","","","","","","","",""})
aAdd (aRegs, {cPerg, "02", "Usuário Com Parâmetros :	", "","", 	"mv_ch2", "C", 30,0,0, 	"G","","mv_par02",""	    	,"","","","",""    			,"","","","",""      	,"","","","","","","","","","","","","","","","USR",""})
aAdd (aRegs, {cPerg, "03", "Usuário Sem Parâmetros :	", "","", 	"mv_ch3", "C", 30,0,0, 	"G","","mv_par03","" 	   		,"","","","",""    			,"","","","",""      	,"","","","","","","","","","","","","","","","USR",""})

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

********************************************************************************
Static Function F1060ValidParam()		// Validações dos campos da pergunta

Local lRet 	:= .F.
Local mv2 	:= AllTrim( mv_par02 )
Local mv3	:= AllTrim( mv_par03 ) 
cLog += dToC(date()) + Space(1) + Time() + Chr(13) + Chr(10) + Chr(13) + Chr(10) + 'Função: '

If ( mv1 == 1 ) 
	If Empty( mv2 ) .or. Empty( mv3 )	
		MsgInfo( 'Ao selecionar a função SUBSTITUIR, os dois campos de usuário devem ser preenchidos.' ) 
	 
	ElseIf MsgYesNo( 'Tem certeza que deseja substituir "' + mv2 + '" por "' + mv3 + '".', 'Substituir' ) 
		cLog += 'Substituir' + Chr(13) + Chr(10) + Chr(13) + Chr(10)
		lRet := .T.
		                                                                                     		
	EndIf
	
ElseIf ( mv1 == 2 )
	If Empty( mv2 ) .or. Empty( mv3 )	
		MsgInfo( 'Ao selecionar a função ESPELHAR, os dois campos de usuário devem ser preenchidos.' )
		
	ElseIf MsgYesNo( 'Tem certeza que deseja espelhar os parâmetros de "' + mv2 + '" para "' + mv3 + '".', 'Espelhar' )
		cLog += 'Espelhar' + Chr(13) + Chr(10) + Chr(13) + Chr(10)
		lRet := .T.          
		
	EndIf
Else
	If Empty( mv2 )
		MsgInfo( 'Ao selecionar a função EXCLUIR, é necessáro informar qual usuário no campo "Usuário Com Parâmetros".' )
				
	ElseIf MsgYesNo( 'Tem certeza que deseja excluir "' + mv2 + '".', 'Excluir' )		   
		cLog += 'Excluir' + Chr(13) + Chr(10) + Chr(13) + Chr(10)
		lRet := .T.
		
	EndIf
EndIf

Return lRet            

********************************************************************************
Static Function F1060PegaUsr()		// Função que captura o Nome e o Id do usuário informado, verificando também se ele está cadastrado no sistema


If IsDigit( Alltrim( mv_par02 ) )
	mv2_id	 :=	AllTrim( mv_par02 )
	mv2_name := ''
	
	PswOrder( 1 )
	
	If !PswSeek( mv2_id, .T. )
		MsgAlert( 'O usuário "' + mv2_id + '" não está cadastrado no Protheus.' )
	    
		Return .F.
	Else
		mv2_name := PswRet()[1][2]
	
	EndIf
	
Else		
	mv2_name := Lower(AllTrim( mv_par02 ))
	mv2_id	:= ''
	
	PswOrder( 2 )
	
	If !PswSeek( mv2_name, .T. )
		MsgAlert( 'O usuário "' + mv2_name + '" não está cadastrado no Protheus.' )
	    
		Return .F.
	Else
		mv2_id	:= PswRet()[1][1]
	
	EndIf                        

EndIf

cLog += 'Usuário com parâmetros: ' + mv2_id + ' - ' + mv2_name + Chr(13) + Chr(10)

If mv1 <> 3		// Se não for uma exclusão pega também o campo Usuário Sem Parâmetros
	If IsDigit( Alltrim( mv_par03 ) )
		mv3_id	 :=	AllTrim( mv_par03 )
		mv3_name := ''
		
		PswOrder( 1 )
		
		If !PswSeek( mv3_id, .T. )
			MsgAlert( 'O usuário "' + mv3_id + '" não está cadastrado no Protheus.' )
		    
			Return .F.
		Else
			mv3_name := PswRet()[1][2]
		
		EndIf
	Else		
		mv3_name := Lower(AllTrim( mv_par03 ))
		mv3_id	 := ''
		
		PswOrder( 2 )
		
		If !PswSeek( mv3_name, .T. )
			MsgAlert( 'O usuário "' + mv3_name + '" não está cadastrado no Protheus.' )
		   
			Return .F.
		Else
			mv3_id	:= PswRet()[1][1]
		
		EndIf                     
	EndIf
	
	cLog += 'Usuário sem parâmetros: ' + mv3_id + ' - ' + mv3_name + Chr(13) + Chr(10)	
	
EndIf 

cLog += Chr(13) + Chr(10)	
                   
Return .T.

********************************************************************************
Static Function F1060PegaDiv( cX6 )		// Pega o caracter separador de usuários
    
Local cDiv	:= ""

If RIGHT( cX6, 1 ) == "/" .or. RIGHT( cX6, 1 ) == ";" .or. RIGHT( cX6, 1 ) == "," .or. RIGHT( cX6, 1 ) == "\" .or. RIGHT( cX6, 1 ) == "|"
	
	cDiv  :=  RIGHT( cX6, 1 )	

ElseIf  "/" $ cX6 .OR. ";" $ cX6 .OR. "," $ cX6 .OR. "\" $ cX6 .OR. "|" $ cX6
    
    If 		"/" $ cX6
    	cDiv 	:= "/" 
    	cDiv1 	:= "/"   
    	
    ElseIf 	";" $ cX6
    	cDiv 	:= ";" 
    	cDiv1 	:= ";"  
    	
    ElseIf 	"," $ cX6
    	cDiv 	:= ","   
    	cDiv1 	:= ","
    	    	
    ElseIf 	"\" $ cX6
    	cDiv 	:= "\"  
    	cDiv1 	:= "\"
    	    	
    ElseIf 	"|" $ cX6
    	cDiv 	:= "|"  
    	cDiv1 	:= "|"
    	    	
    EndIf               
    
ElseIf "@slcalimentos" $ cX6 	
	
	cDiv 	:= ";" 
	cDiv1	:= ";"
	
Else
	cDiv 	:= ","	  
	cDiv1	:= ","
	
EndIf

Return cDiv   

********************************************************************************
Static Function F1060Grava()    

lGOk := .F.
For i := 1 To Len( aList )
  
	If aList[i][1] .AND. Len( aList[i][9] ) <= 250	// Se estiver selecionado e o tamanho do X6_CONTEUD couber o novo usuário
	
		If SX6->( dbSeek( aList[i][2] + aList[i][3] ) )
		
	        RecLock("SX6", .F.) 
	        
	        	SX6->X6_CONTEUD := aList[i][9]
	        
	        MsUnlock()            
            
		MsgInfo( 'Parâmetros atualizados com sucesso!' )
	        lGOk := .T.	               
	        cLog  +=  aList[i][2] + ' + ' + aList[i][3] + ' -> ' + aList[i][9] + Chr(13) + Chr(10)
	                    
	Else
	    	MsgAlert( 'Não foi possível se posicionar na posição ' + aList[i][2] + ' + ' + aList[i][3] )        
	        cLog  +=  aList[i][2] + ' + ' + aList[i][3] + ' -> ERRO! Não foi possível se posicionar neste parâmetro.' + Chr(13) + Chr(10)
	        
	EndIf
	ElseIf  Len( aList[i][9] ) > 250
	
	    MsgAlert( 'ERRO! O novo conteúdo excede o tamanho do campo' )
	    lGOk := .F.
	    cLog  +=  aList[i][2] + ' + ' + aList[i][3] + ' -> ERRO! O novo conteúdo excede o tamanho do campo.' + Chr(13) + Chr(10)		
	EndIf   
Next  
  
Return                

********************************************************************************
Static Function F1060MontaTela()        

Local lOk
//Define Font oFntGet1 Name 'Calibri' Size 30,50 Bold

aAdd( aButPed, { 'PMSINFO', {|| SelectAll( aList ) }, 'Marcar/Desmarcar Todos' } )     // Adicionado um botão no submenu 'Ações Relacionadas'

aSize	 := MsAdvSize()
aObjects := {}

AAdd( aObjects, { 100, 100, .t., .t. } )
AAdd( aObjects, { 100, 030, .t., .f. } )
aInfo 	:= { aSize[ 1 ], aSize[ 2 ], aSize[ 3 ], aSize[ 4 ], 3, 3 }
aPosObj := MsObjSize( aInfo, aObjects )

Define msDialog oDlg1 Title 'Seleção de Parâmetros' From aSize[7],0 to aSize[6],aSize[5] OF oMainWnd PIXEL

@ aPosObj[1,1]+000,aPosObj[1,2]  ListBox oList ; //45,05
Fields Header  '     ' ,'Filial' , 'Variável', 'Tipo', 'Descrição', 'Conteúdo Atual' ;
Size aPosObj[1,4]-aPosObj[1,2],aPosObj[1,3]-aPosObj[1,1] OF oDlg1 PIXEL ColSizes 60,30 ;
Pixel Of oDlg1 ;
On dblClick( aList:=SelectBox( oList:nAt, aList ), oList:Refresh() )


oList:aColSizes := {5,		  15,	   25,		 15,	   330,					  330   		}
		//  ListBox 	| X6_FIL | X6_VAR      | X6_TIPO | X6_DESCRIC + X6_DESC1 + X6_DESC2     | X6_CONTEUD ( cString )
				
oList:SetArray( aList )

oList:bHeaderClick := { |oObj,nCol| If( nCol==1, SelectAll( aList ), Nil), oList:Refresh() }

oList:bLine := { || { ;	
	Iif( aList[ oList:nAt, 01 ], oOk, oNo ), ; 		// ListBox
	AllTrim( aList[ oList:nAt, 02 ] ),  ;			// X6_FIL                         
	AllTrim( aList[ oList:nAt, 03 ] ),  ;			// X6_VAR
	AllTrim( aList[ oList:nAt, 04 ] ),  ;			// X6_TIPO   
	;// X6_DESCRIC + X6_DESC1 + X6_DESC2 => Monta a descrição sem repetição
	Iif( AllTrim(aList[ oList:nAt, 05 ]) != AllTrim(aList[ oList:nAt, 06 ]), ;
		Iif ( AllTrim(aList[ oList:nAt, 06 ]) != AllTrim(aList[ oList:nAt, 07 ]), ;
			AllTrim(aList[ oList:nAt, 05 ]) + ' ' + AllTrim(aList[ oList:nAt, 06 ]) + ' ' + AllTrim(aList[ oList:nAt, 07 ]), ;
			AllTrim(aList[ oList:nAt, 05 ]) + ' ' + AllTrim(aList[ oList:nAt, 06 ]) ;
		), ;
		AllTrim(aList[ oList:nAt, 05 ]) ;
	), ;
	AllTrim( aList[ oList:nAt, 08 ] )   ;			// X6_CONTEUD
}}

oList:nAt:=1
oList:Refresh()

Activate msDialog oDlg1 Centered On Init EnchoiceBar( oDlg1, { || If ( Confirm( aList ), lOk := .t., lOk := .f. ) , If ( lOk == .t., oDlg1:End(), Nil ) }, { || aList := Array( 0 ), oDlg1:End() },, aButPed)                                                                                                                                                                      		

GetDRefresh() 

Return lOk    

********************************************************************************
Static Function SelectBox( nIt, aVector )

If !aVector[ nIt, 1 ]
	aVector[ nIt, 1 ] := .t.	
Else	                 
	aVector[ nIt, 1 ] := .f.
EndIf
    
oList:Refresh()
Return( aVector ) 

********************************************************************************
Static Function SelectAll( aVector )

For i := 1 To Len(aVector)
	If !aVector[ i, 1 ]
		aVector[ i, 1 ] := .t.
	Else
		aVector[ i, 1 ] := .f.   	
	EndIf
Next i

oList:Refresh()

Return( aVector )       

********************************************************************************
Static Function Confirm ( aVector )

Local lControle  :=  .F.

For i := 1 To Len(aVector)
	If aVector[i][1]  ==  .T.  // se está marcado
		lControle  :=  .T.
		Exit
	EndIf
Next

If !lControle       
	MsgInfo("Você não selecionou nenhum parâmetro.")	
EndIf        

Return lControle 
