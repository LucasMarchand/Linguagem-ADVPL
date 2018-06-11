#include 'rwmake.ch' 
#include 'protheus.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCF1060      || Autor: Lucas Rocha          || Data: 05/12/17  ||
||-------------------------------------------------------------------------||
|| Descrição: Ajuste de usuários nos parâmetros da SX6			   ||		                                   
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/                                            

User Function SLCF1060()

//////////////////////////////////
// 	    Variáveis  		//
//////////////////////////////////
Local cPerg	:= 'SLCF1060'
Local lCond 	:= ''
Local cString   := ''
Local cX6	:= ''
Local cDiv	:= ''
Local cDir    	:= '\logs\'
Local cArq    	:= 'LogSX6_' 
Local nHandle 	:= 0
Local lAltera	:= .F.

Private mv1 	 := 0
Private mv2_name := Space(30)
Private mv3_name := Space(30)
Private mv2_id	 := Space(6)
Private mv3_id	 := Space(6)
Private cLog	 := '' 
Private aList	 := Array(0)   
Private oDlg1
Private nMsg	 := 0
Private oList
Private oOk	 := Loadbitmap( GetResources(), 'LBOK' )
Private oNo	 := Loadbitmap( GetResources(), 'LBNO' )
Private aButPed  := Array(0) 
Private cFunc
Private cParams 
Private lDivApos
Private lDivFim

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
	lAltera	 := .T.
	lDivApos := .F.
	lDivFim  := .F.                                                                       
	cString  := ''
	cX6 	 := AllTrim( X6_CONTEUD )
	cDiv 	 := F1060PegaDiv( cX6, mv2_name )	// Pega o caractere que separa os usuários   
            
	If ( mv1 == 1 ) .AND. !( mv3_name $ cX6 )		// Substituir	   
		cString := StrTran( cX6, mv2_name, mv3_name )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_name $ cX6 ) 	// Espelhar
	
		If SUBSTR( cX6, AT(mv2_name, cX6) + LEN(mv2_name), LEN('@slcalimentos') ) == '@slcalimentos'   // Se estiver gravado como e-mail	    		    	

			If lDivFim
				cString := cX6 + mv3_name + '@slcalimentos.com.br' + cDiv

			Else
				cString := cX6 + cDiv + mv3_name + '@slcalimentos.com.br' + cDiv  

			EndIf 	    	

	 	ElseIf lDivFim		
			cString := cX6 + mv3_name + cDiv 
		
		Else
			cString := cX6 + cDiv + mv3_name + cDiv
					       
		EndIf
											
	ElseIf ( mv1 == 3 )		// Excluir    

		If SUBSTR( cX6, AT(mv2_name, cX6) + LEN(mv2_name) , LEN('@slcalimentos') ) == '@slcalimentos' 						
			
			If lDivApos			
				cString := StrTran( cX6, mv2_name + '@slcalimentos.com.br' + cDiv, '' ) 	// Se possuir divisor irá excluir 
		   
			Else
				cString := StrTran( cX6, mv2_name + '@slcalimentos.com.br', '' ) 
				
			EndIf
		Else  
			If lDivApos	
				cString := StrTran( cX6, mv2_name + cDiv, '' )  
			
			Else
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
	lAltera	 := .T.
	lDivApos := .F.
	lDivFim  := .F.
	cString  := ''
	cX6 	 := AllTrim( X6_CONTEUD )	
	cDiv 	 := F1060PegaDiv( cX6, mv2_id )		// Pega o caractere que separa os usuários
         
	If ( mv1 == 1 ) .AND. !( mv3_id $ cX6 )		// Substituir
		
		cString := StrTran( cX6, mv2_id, mv3_id )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_id $ cX6 ) // Espelhar
		
		If lDivFim
			cString := cX6 + mv3_id + cDiv
			
		Else
			cString := cX6 + cDiv + mv3_id + cDiv
		
		EndIf			       											
	ElseIf ( mv1 == 3 )							// Excluir		
		
		If lDivApos
			
			cString := StrTran( cX6, mv2_id + cDiv, '' )
		Else                                          
		
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

// Criar tela
If Len(aList) == 0
   
	MsgInfo( 'Não existem parâmetros para serem configurados.' )   

	Return U_SLCF1060()	
	
Else	
	If F1060MontaTela() 		
		MSAguarde ({ || F1060Grava() }, "Aguarde", "Gravando dados no banco..." ) 
				
		// Salvar cLog ==> Na pasta '\TOTVS 12\Microsiga\Protheus_Data\logs\LogSX6_Usuário_Data_Hora' 
		if mv1 == 3
			cArq += StrTran(mv2_name, '.', '-')
			
		Else
			cArq += StrTran(mv3_name, '.', '-')
		EndIf                       
		
		cArq += '_' + DTOS( DATE() ) + '_' + SUBSTR(TIME(), 1, 2) + SUBSTR(TIME(), 4, 2) + SUBSTR(TIME(), 7, 2) + '.txt'
		nHandle := FCreate( cDir + cArq )
		
		If nHandle < 0
			MsgAlert( 'Erro durante a gravação do log. Verifique se existe o caminho: ' + cDir + cArq  )
			
		Else
			FWrite( nHandle, cLog )
			FClose( nHandle )
			MsgInfo( 'Log salvo no diretório "Protheus_Data' + cDir + cArq + '"' ) 
			ShellExecute('open', cArq, '', '\\Alccosrpt01\TOTVS 12\Microsiga\Protheus_Data\logs\', 1)
		EndIf   		
	Else
		Return U_SLCF1060()
		
	EndIf
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
aAdd (aRegs, {cPerg, "01", "Função				", "","", 	"mv_ch1", "N",  1,0,0, 	"C","","mv_par01","Substituir","Substituir","Substituir",""		,"","Espelhar","Espelhar","Espelhar",""		,"","Excluir","Excluir","Excluir","","","","","","","","","","","","","","",""})
aAdd (aRegs, {cPerg, "02", "Usuário parametrizado :		", "","", 	"mv_ch2", "C", 30,0,0, 	"G","","mv_par02","","","",""	,"","","","",""		,"","","","",""		,"","","","",""		,"","","","",""		,"USR","","",""})
aAdd (aRegs, {cPerg, "03", "Usuário não-parametrizado :		", "","", 	"mv_ch3", "C", 30,0,0, 	"G","","mv_par03","","","",""	,"","","","",""		,"","","","",""		,"","","","",""		,"","","","",""		,"USR","","",""})

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
	Else
		cFunc := 'Substituir'
		lRet := .T.
	EndIf
	
ElseIf ( mv1 == 2 )
	If Empty( mv2 ) .or. Empty( mv3 )	
		MsgInfo( 'Ao selecionar a função ESPELHAR, os dois campos de usuário devem ser preenchidos.' )		
	Else
		cFunc := 'Espelhar'
		lRet := .T.		
	EndIf 
	
Else
	If Empty( mv2 )
		MsgInfo( 'Ao selecionar a função EXCLUIR, é necessáro informar qual usuário no campo "Usuário Com Parâmetros".' )
	Else
		cFunc := 'Excluir'
		lRet := .T.		
	EndIf
EndIf

cLog += cFunc + Chr(13) + Chr(10) + Chr(13) + Chr(10)

Return lRet            

********************************************************************************
Static Function F1060PegaUsr()		// Função que captura o Nome e o Id do usuário informado, verificando também se ele está cadastrado no sistema


If IsDigit( Alltrim( mv_par02 ) )
	mv2_id	 := AllTrim( mv_par02 )
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

cLog += 'Usuário parametrizado: ' + mv2_id + ' - ' + mv2_name + Chr(13) + Chr(10)

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
	
	cLog += 'Usuário não-parametrizado: ' + mv3_id + ' - ' + mv3_name + Chr(13) + Chr(10)
	
EndIf 

cLog += Chr(13) + Chr(10)	
                   
Return .T.

********************************************************************************
Static Function F1060PegaDiv( cX6, mv2 )		// Pega o caracter separador de usuários
    
Local cDiv	:= ""
Local cTemp, cTemp1

cTemp 	:= SUBSTR( cX6, AT(mv2, cX6) + LEN(mv2) , 1 )                               // Posição do possível caractere depois do usuário
cTemp1	:= SUBSTR( cX6, AT(mv2, cX6) + LEN(mv2) + LEN("@slcalimentos.com.br"), 1 )  // Posição do possível caractere depois do usuário, quando é escrito como e-mail

If RIGHT( cX6, 1 ) $ "/;,\|"
	
	lDivFim := .T.
EndIf

If cTemp $ "/;,\|"
	cDiv 	 := cTemp 
	lDivApos := .T.

ElseIf cTemp1 $ "/;,\|"	
	cDiv  :=  cTemp1
	lDivApos  := .T.
    
ElseIf 	"/" $ cX6
	cDiv 	:= "/"   

ElseIf 	"," $ cX6
	cDiv 	:= "," 
    	
ElseIf 	";" $ cX6  .OR. "@slcalimentos" $ cX6
	cDiv 	:= ";"   
    	    	
ElseIf 	"\" $ cX6
	cDiv 	:= "\"  
    	    	
ElseIf 	"|" $ cX6
	cDiv 	:= "|"  
	
Else
	cDiv 	:= ","	  
	
EndIf

Return cDiv             

********************************************************************************
Static Function F1060MontaTela()        

Local lOk
Define Font oFntGet1 Name 'TAHOMA' Size 10,15 Bold
Define Font oFntGet2 Name 'TAHOMA' Size 10,15

aAdd( aButPed, { 'PMSINFO', {|| SelectAll( aList ) }, 'Marcar/Desmarcar Todos' } )     // Adicionado um botão no submenu 'Ações Relacionadas'

aSize	 := MsAdvSize()
/*
aSize = {	1=Linha inicial área trabalho,
		2=Coluna inicial área trabalho,
		3=Linha final área trabalho,
		4=Coluna final área trabalho,
		5=Coluna final dialog (janela),
		6=Linha final dialog (janela),
		7=Linha inicial dialog (janela) }*/

// Agora montamos uma array com os elementos da tela:
// Aonde devemos informar
// AAdd( aObjects, { Tamanho X (horizontal) , Tamanho Y (vertical), Dimensiona X , Dimensiona Y, Retorna dimensões X e Y ao invés de linha / coluna final } )
aObjects := {}
AAdd( aObjects, { 100, 030, .t., .f. } )
AAdd( aObjects, { 100, 100, .t., .t. } )

// Montamos a array com o valor da tela, aonde:
// aInfo := { 1=Linha inicial, 2=Coluna Inicial, 3=Linha Final, 4=Coluna Final, Separação X, Separação Y, Separação X da borda (Opcional), Separação Y da borda (Opcional) }
aInfo := { aSize[ 1 ], aSize[ 2 ], aSize[ 3 ], aSize[ 4 ], 5, 5, 5, 5 }

// Passamos agora todas as informações para o calculo das dimenções:
// MsObjSize( aInfo, aObjects, Mantem Proporção , Disposição Horizontal )
aPosObj := MsObjSize( aInfo, aObjects )

aSort(aList, , , { | x,y | x[2] + x[3] < y[2] + y[3] })  // Ordena por X6_FIl + X6_VAR

Define msDialog oDlg1 Title 'Seleção de Parâmetros' From aSize[7],0 to aSize[6],aSize[5] OF oMainWnd PIXEL

@ aPosObj[1,1]+001,aPosObj[1,2] Say 'Função: ' Font oFntGet2 Size 300,300 Pixel Of oDlg1
@ aPosObj[1,1]+001,aPosObj[1,2]+033 Say oFunc Var cFunc Font oFntGet1 Size 300,300 Pixel Colors CLR_HRED Of oDlg1 

@ aPosObj[1,1]+012,110 Say oPessoa1 Var mv2_name Font oFntGet2 Size 300,300 Pixel Of oDlg1
If mv1 != 3
	@ aPosObj[1,1]+012,200 Say ' -> ' Font oFntGet2 Size 300,300 Pixel Of oDlg1
	@ aPosObj[1,1]+012,235 Say oPessoa2 Var mv3_name Font oFntGet2 Size 300,300 Pixel Of oDlg1
EndIf

@ aPosObj[2,1]+000,aPosObj[2,2]  ListBox oList ; //45,05
Fields Header  '     ' ,'Filial' , 'Variável', 'Tipo', 'Descrição', 'Conteúdo Atual' ;
Size aPosObj[2,4]-aPosObj[2,2],aPosObj[2,3]-aPosObj[2,1] OF oDlg1 PIXEL ColSizes 60,30 ; //555,148
Pixel Of oDlg1 ;
On dblClick( aList:=SelectBox( oList:nAt, aList ), oList:Refresh() )


oList:aColSizes := {15,	      15,      25	 15,	   400,				      500 }
		//  ListBox | X6_FIL | X6_VAR  | X6_TIPO | X6_DESCRIC + X6_DESC1 + X6_DESC2 | X6_CONTEUD ( cString )
				
oList:SetArray( aList )

oList:bHeaderClick := { |oObj,nCol| If( nCol==1, SelectAll( aList ), Nil), oList:Refresh() }

oList:bLine := { || { ;	
	Iif( aList[ oList:nAt, 01 ], oOk, oNo ), ; 		// ListBox
	AllTrim( aList[ oList:nAt, 02 ] ),  ;			// X6_FIL                         
	AllTrim( aList[ oList:nAt, 03 ] ),  ;			// X6_VAR
	AllTrim( aList[ oList:nAt, 04 ] ),  ;			// X6_TIPO   
	;	// X6_DESCRIC + X6_DESC1 + X6_DESC2 => Monta a descrição sem repetição
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
cParams := ''

For i := 1 To Len(aVector)
	If aVector[i][1]  ==  .T.  // se está marcado
		lControle  :=  .T.

		If Empty(cParams) 
			cParams := AllTrim( aVector[i][3] )
		Else
			cParams := cParams + ', ' + AllTrim( aVector[i][3] )
		EndIf
	EndIf
Next i

If !lControle
	MsgInfo("Você não selecionou nenhum parâmetro.")
	Return .F.	
EndIf

If MessageBox("Você tem certeza que deseja " + Lower(cFunc) + " estes parâmetros: " + cParams + " ?", Upper(cFunc), 4) <> 6
	Return .F.
EndIf

Return .T.

********************************************************************************
Static Function F1060Grava()    

Local cFilVar

For i := 1 To Len( aList )

	cFilVar := IIF( Empty(aList[i][2]), '     ' + aList[i][3], AllTrim(aList[i][2]) + ' + ' + aList[i][3] )
  
	If aList[i][1]  // Se estiver marcado
		If Len( aList[i][9] ) <= 250	// Se o tamanho do X6_CONTEUD couber o novo usuário
		
			If SX6->( dbSeek( aList[i][2] + aList[i][3] ) )
			
		        	RecLock("SX6", .F.) 
		        
		        	SX6->X6_CONTEUD := aList[i][9]
		        
		        	MsUnlock()            
	            
				cLog  += cFilVar + ' -> ' + aList[i][9] + Chr(13) + Chr(10) 
	
			Else
				MsgAlert( 'Não foi possível se posicionar na posição "' + aList[i][2] + '" + "' + aList[i][3] + '"! ' )        
				cLog  +=  cFilVar + ' -> ERRO! Não foi possível se posicionar neste parâmetro.' + Chr(13) + Chr(10)
		        
			EndIf
		ElseIf  Len( aList[i][9] ) > 250
		
			MsgAlert( 'ERRO! O novo conteúdo excede o tamanho do campo ' + AllTrim(cFilVar) )
			cLog  +=  cFilVar + ' -> ERRO! O novo conteúdo excede o tamanho do campo.' + Chr(13) + Chr(10)		
		EndIf
	EndIf   
Next  

SX6->( dbCloseArea() )

Return
