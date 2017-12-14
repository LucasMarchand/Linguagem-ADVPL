#include 'rwmake.ch'

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
// 	    Variáveis  		//
//////////////////////////////////
Local cPerg	:= 'SLCF1060'
Local lCond 	:= ''
Local lCond1	:= ''
Local lMsg	:= .F.
Local lMsg1	:= .F.
Local cString   := ''
Local cX6	:= ''
Local cDiv	:= ''
Local nOpc	:= 0
Local cDir    	:= '\logs\'
Local cArq    	:= 'LogSX6_' + DTOS( DATE() )  + '_' + SUBSTR(TIME(), 1, 2) + SUBSTR(TIME(), 4, 2) + SUBSTR(TIME(), 7, 2) + '.txt' 
Local nHandle 	:= 0

Private mv1	 := 0
Private mv2_name := Space(30)
Private mv3_name := Space(30)
Private mv2_id	 := Space(6)
Private mv3_id	 := Space(6)
Private cDiv1  	 := '' 
Private cLog	 := '' 

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
dbSetOrder(1)  

// Pesquisa pelo NOME do usuário
lCond	:= ' "' + mv2_name + '" $ X6_CONTEUD '
    
SX6->( dbSetFilter( { || &lCond }, lCond) )

Count TO nLinha		
If nLinha == 0
	lMsg := .T.
EndIf    

dbGoTop()

// Se encontrou algum varre 
WHILE !EOF()
	cX6 	:= AllTrim( X6_CONTEUD )
	cDiv1 	:= ''
	cDiv 	:= F1060PegaDiv( cX6 )	// Pega o caractere que separa os usuários
    
         
	If ( mv1 == 1 ) .AND. !( mv3_name $ cX6 )		// Substituir
	   
		cString := StrTran( cX6, mv2_name, mv3_name )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_name $ cX6 ) 	// Espelhar
	
		If SUBSTR( cX6, AT(mv2_name, cX6) + LEN(mv2_name), LEN('@slcalimentos') ) == '@slcalimentos'   // Se estiver gravado como e-mail	    	
	    	cString := cX6 + cDiv1 + mv3_name + '@slcalimentos.com.br' + cDiv  
	    	
	 	Else		
			cString := cX6 + cDiv1 + mv3_name + cDiv 
						       
		EndIf
											
	Else		// Excluir    

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
	EndIf		
	
	F1060Grava( cString )
	 	 
	dbSkip()
END

SX6->( dbClearFilter() )
 

// Pesquisa pelo ID do usuário
lCond	:= ' "' + mv2_id + '" $ X6_CONTEUD '
  
SX6->( dbSetFilter( { || &lCond }, lCond) )

Count TO nLinha1		
If nLinha1 == 0
	lMsg1 := .T.
EndIf    

dbGoTop()

// Varre todos parâmetros achados 
WHILE !EOF()
	cX6 	:= AllTrim( X6_CONTEUD )
	cDiv1 	:= ''	
	cDiv 	:= F1060PegaDiv( cX6 )		// Pega o caractere que separa os usuários
    
         
	If ( mv1 == 1 ) .AND. !( mv3_id $ cX6 )		// Substituir
		cString := StrTran( cX6, mv2_id, mv3_id )			       
	
	ElseIf ( mv1 == 2 ) .AND. !( mv3_id $ cX6 ) // Espelhar
		cString := cX6 + cDiv1 + mv3_id + cDiv			       
											
	Else 										// Excluir
		cString := StrTran( cX6, mv2_id + cDiv, '' )
		
		If cX6 == cString
			 cString := StrTran( cX6, mv2_id, '' )
			
		EndIf
		
	EndIf		
	
	F1060Grava( cString )	 	
 
	dbSkip()
END                     

SX6->( dbClearFilter() )
SX6->( dbCloseArea() )
 

If lMsg .AND. lMsg1 		
	MsgInfo( 'Não existem parâmetros registrados para este usuário!' )   

	Return U_SLCF1060() 
Else
	MsgInfo( 'Configuração realizada com sucesso!' ) 
	
	// Salvar cLog ==> Na pasta '\TOTVS 11\Microsiga\Protheus_Data\logs\LogSX6_Data_Hora'
	nHandle := FCreate( cDir + cArq )
	
	If nHandle < 0
		MsgAlert( 'Erro durante a gravação do log.' )
		
	Else
		FWrite( nHandle, cLog )
		FClose( nHandle )
		MsgInfo( 'Log salvo no diretório "Protheus_Data' + cDir + cArq + '"' ) 
		
	EndIf		   
EndIf

Return

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
Static Function F1060ValidParam()		// Validações dos campos de pergunta

Local lRet 	:= .F.
Local mv2 	:= AllTrim( mv_par02 )
Local mv3	:= AllTrim( mv_par03 ) 
cLog  +=  dToC(date()) + Space(1) + Time() + Chr(13) + Chr(10) + Chr(13) + Chr(10) + 'Função: '

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
	
	PswOrder( 1 )
	
	If !PswSeek( mv2_id, .T. )
		MsgAlert( 'O usuário "' + mv2_id + '" não está cadastrado no Protheus.' )
	    
		Return .F.
	Else
		mv2_name := PswRet()[1][2]
	
	EndIf
Else		
	mv2_name := AllTrim( mv_par02 )
	
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
		
		PswOrder( 1 )
		
		If !PswSeek( mv3_id, .T. )
			MsgAlert( 'O usuário "' + mv3_id + '" não está cadastrado no Protheus.' )
		    
			Return .F.
		Else
			mv3_name := PswRet()[1][2]
		
		EndIf
	Else		
		mv3_name := AllTrim( mv_par03 )
		
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
Static Function F1060PegaDiv( cX6 )
    
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
Static Function F1060Grava( cString )    

If LEN( cString ) <= 250

	If SX6->( dbSeek( X6_FIL + X6_VAR ) )
	
        RecLock("SX6", .F.) 
        
        	SX6->X6_CONTEUD := cString
        
        MsUnlock()            
        
        // MsgInfo( 'Parâmetro alterado --> ' + X6_VAR )        
        cLog  +=  X6_FIL + ' + ' + X6_VAR + ' -> ' + cString + Chr(13) + Chr(10)
        
        Return .T.
                    
    Else
    	MsgAlert( 'Não foi possível se posicionar na posição ' + X6_FIL + ' + ' + X6_VAR )        
        cLog  +=  X6_FIL + ' + ' + X6_VAR + ' -> ERRO! Não foi possível se posicionar neste parâmetro.' + Chr(13) + Chr(10)
        
	EndIf
Else
    MsgAlert( 'ERRO! O novo conteúdo excede o tamanho do campo' )
    cLog  +=  X6_FIL + ' + ' + X6_VAR + ' -> ERRO! O novo conteúdo excede o tamanho do campo.' + Chr(13) + Chr(10)	

EndIf   

Return .F.
