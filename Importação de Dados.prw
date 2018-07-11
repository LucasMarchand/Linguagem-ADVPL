#include 'fivewin.ch'
#Include 'rwmake.ch'
#include 'tbiconn.ch'  
#Include 'Protheus.ch'   

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCA990      || Autor: Lucas Rocha           || Data: 13/04/18  ||
||-------------------------------------------------------------------------||
|| Descrição: Rotina para importação dos dados da Target                   ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||*/

User Function SLCA990()

Private aDados		:= {}
Private cNovoArq
Private cPerg		:= 'SLCA990'

Private cCadastro	:= ''
					
Private aRotina	:= {{ 'Visualizar'  , 'AxVisual'  			, 0, 2 }, ; 
				    { 'Alterar'     , 'AxAltera'		    , 0, 4 }, ;
				    { 'Excluir'     , 'AxDeleta'		    , 0, 5 }, ;
				    { 'Importa'     , 'processa( {|| u_A990_ImportaArq() }, "Importando arquivo", "Aguarde...", .f. ) ' , 0, 3 }, ;
            		{ 'Relatório'   , 'u_SLCR990()'	        , 0, 1 } }
//          		{ 'Legenda'     , 'u_A990_Legenda()'	, 0, 6 } }

// Private aCores	:= {{ "Z69_STATUS=='0'",    'BR_BRANCO'         }, ;
//                     { "Z69_STATUS=='1'",    'BR_AZUL'           }, ;   
//                     { "Z69_STATUS=='2'",    'BR_AZUL_CLARO'     }, ;
//                     { "Z69_STATUS=='3'",    'BR_PINK'           }, ;
//                     { "Z69_STATUS=='4'",    'BR_VERDE'          }, ;
//                     { "Z69_STATUS=='5'",    'BR_VERDE_ESCURO'   }, ;
//                     { "Z69_STATUS=='6'",    'BR_MARRON_OCEAN'   }, ;
//                     { "Z69_STATUS=='7'",    'BR_LARANJA'        }  } 

dbSelectArea( 'Z69' )
dbSetOrder( 1 )
dbGoTop()

mBrowse( 06, 01, 22, 75, 'Z69',,,,,, /*aCores*/ )

Return

*********************************************************************************************************
User Function A990_Legenda()

/*
0 - Possui na Target, mas não no Protheus
1 - Chaves NFE iguais
2 - Chave NFE diferentes / Chave Filial + Numero + Fornecedor + Loja
3 - Nota da SLC
4 - Total igual
5 - Total igual / Falta classificação
6 - NF Produtor
7 - Nota de terceiros
*/

Local aLegenda := {}

aAdd( aLegenda, {"BR_BRANCO",	    ""	})	//Z69_STATUS == '0'
aAdd( aLegenda, {"BR_AZUL",	        ""	})	//Z69_STATUS == '1'
aAdd( aLegenda, {"BR_AZUL_CLARO",   ""	})	//Z69_STATUS == '2'
aAdd( aLegenda, {"BR_PINK",	        ""	})	//Z69_STATUS == '3'
aAdd( aLegenda, {"BR_VERDE",	    ""	})	//Z69_STATUS == '4'
aAdd( aLegenda, {"BR_VERDE_ESCURO", ""	})	//Z69_STATUS == '5'
aAdd( aLegenda, {"BR_MARRON_OCEAN",	""	})	//Z69_STATUS == '6'
aAdd( aLegenda, {"BR_LARANJA",	    ""	})	//Z69_STATUS == '7'

BrwLegenda( "", "Protheus x Target", aLegenda )

Return

*********************************************************************************************************
Static Function A990_LeiaArq()

nHandle := FOpen( cNovoArq )

If nHandle < 0
	MsgAlert( 'Erro durante a abertura do arquivo!' )
	Return .F.
Endif

FT_FUse( cNovoArq )
FT_FGoTop()
 
While !FT_FEOF()

	If !Empty( FT_FREADLN() )
		aadd( aDados, Separa( FT_FREADLN(), ";", .T. ) )   // Separa dentro do array aDados: Filial [1] - Número [2] - Fornecedor [3] - Loja [4]
		FT_FSKIP()
	EndIf
	
EndDo
 
FT_FUse()
FClose( nHandle )

Return          

*********************************************************************************************************
Static Function ValidPerg( cPerg )  	// Cria o grupo de perguntas

Local aArea    := GetArea()                                                                              
Local aAreaSX1 := SX1->(GetArea())

dbSelectArea("SX1")
dbSetOrder(1)
cPerg := Padr(cPerg,10)
aRegs:={}
                                       
aAdd (aRegs, {cPerg, "01", "Arquivo	:			", "","", 	"mv_ch1", "C", 99,0,0, 	"G","U_A990_ProcuraArq()","mv_par01","","","",""	,"","","","",""		,"","","","",""		,"","","","",""		,"","","","",""		,"","","",""})

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

*********************************************************************************************************
User Function A990_ProcuraArq()

cNovoArq  := cGetFile( 'Arquivo Texto (*.TXT) |*.TXT |', 'Selecione o arquivo *.TXT', 0,, .t., GETF_OVERWRITEPROMPT + GETF_LOCALFLOPPY + GETF_LOCALHARD + GETF_NETWORKDRIVE )

If !Empty( cNovoArq )	
	&( ReadVar() ) := PadR( cNovoArq, Len( &( ReadVar() ) ) )
EndIf

Return

*********************************************************************************************************
User Function A990_ImportaArq()

Local aArea     := GetArea()
Local nCount    := 0
Local aFilial   := {}
Local cForne, cLoja, nPos, nI
Local cLog      := ''
Local cLogAux   := ''

ValidPerg( cPerg )
If !Pergunte( cPerg, .T. )
	Return
EndIf

If Empty( cNovoArq )
  MsgInfo('Selecione novamente o arquivo a ser aberto.')
  Return u_A990_ImportaArq()                                                     
Else	
	A990_LeiaArq()
EndIf    

ProcRegua( Len(aDados) )

// Revisar cadastro de cnpj
aAdd( aFilial, {'SLC MATRIZ',			'01',   '04107020000117' } )
aAdd( aFilial, {'SLC  - ALEGRETE',		'02',   '04107020000206' } )
aAdd( aFilial, {'SLC - TATUI',			'07',   '04107020000389' } )
aAdd( aFilial, {'SLC - BAHIA',			'08',   '04107020000893' } )
aAdd( aFilial, {'SLC - PERNAMBUCO',		'10',   '04107020001008' } ) // dupl
aAdd( aFilial, {'SLC JANDIRA',			'11',   '04107020001199' } )
aAdd( aFilial, {'SLC - FORTALEZA',		'13',   '04107020001350' } )
aAdd( aFilial, {'SLC - TOCANTINS',		'15',   '04107020001512' } )
aAdd( aFilial, {'SLC - CAPÃO DO LEÃO',	'17',   '04107020001784' } )
aAdd( aFilial, {'SLC CONCEIÇÃO', 		'20',   '04107020002080' } )
aAdd( aFilial, {'SLC - RECIFE - PE',	'22',   '04107020001008' } ) // dupl
aAdd( aFilial, {'SLC – FORMOSA – GO',	'23',   '04107020002322' } )

dbSelectArea('SA1')
dbSelectArea('SA2')
SA1->( DBSetOrder(3) )
SA2->( DBSetOrder(3) )

For nI := 1 to Len( aDados )

    If Empty( aDados[ nI, 1 ] )
        Loop
    EndIf

    cForne  := ''
    cLoja   := ''
	
    If !Empty( aDados[nI, 3] )       
        If SA2->(dbSeek( xFilial('SA2') + AllTrim( aDados[nI, 3] ), .F. ))
			cForne	:=	SA2->A2_COD
			cLoja	:=	SA2->A2_LOJA
		ElseIf SA1->(dbSeek( xFilial('SA1') + AllTrim( aDados[nI, 3] ), .F. ))      
			cForne	:=	SA1->A1_COD
			cLoja	:=	SA1->A1_LOJA
		EndIf
    Endif
    
    // Verifica se o campo da chave possui 44 carcteres e se refere mesmo a uma NF
    If len( AllTrim( aDados[nI, 4] ) ) == 44 .AND. Substr( AllTrim( aDados[nI, 4] ), 21, 2 ) == '55'

        nPos    	:= aScan( aFilial, { |x| x[1] == aDados[nI, 1] } )
        //lTransf 	:= aScan( aFilial, { |x| x[3] == aDados[nI, 3] } )    // Quando for transferência de filial

        If nPos == 0
            cLogAux += 'li ' + cValToChar( nLinha ) + ' - Não encontrada a filial do campo.' + Chr(13) + Chr(10) 
            Loop
        EndIf

        dbSelectArea( 'Z69' )

        Z69->( dbSetOrder(1) )  // Filial + Numero + Série + Fornecedor + Loja
        If dbSeek( aFilial[nPos][2] + StrZero( Val( aDados[nI, 9]) , 9 ) + PADR( aDados[nI, 8], 3 ) + cForne + cLoja, .F. ) 
            Reclock( 'Z69', .F. )
            Z69->( DbDelete() )
            MsUnlock()
        EndIf

        Z69->( dbSetOrder(2) )  // Filial + Chave
        If dbSeek( aFilial[nPos][2] + AllTrim( aDados[nI, 4] ), .F. )
            Reclock( 'Z69', .F. )
            Z69->( DbDelete() )
            MsUnlock()
        EndIf

        Reclock( 'Z69', .T. )

        Z69->Z69_FILIAL :=	aFilial[ nPos ][ 2 ]
        Z69->Z69_NUM	:=	StrZero( Val(aDados[nI, 9 ]), 9 )
        Z69->Z69_SERIE	:=	PadR( aDados[nI, 8], 3 )
        Z69->Z69_FORNE	:=	cForne
        Z69->Z69_LOJA	:=	cLoja
        Z69->Z69_CHAVE	:=	AllTrim(aDados[nI, 4])
        Z69->Z69_CFOP   :=  AllTrim(aDados[nI, 5])
        Z69->Z69_NOME   :=  AllTrim(aDados[nI, 6])  // razao social
        Z69->Z69_EMIS   :=  cToD(aDados[nI, 7])
        Z69->Z69_CGC    :=  AllTrim(aDados[nI, 3])
        Z69->Z69_TOTAL  :=  Val(StrTran(StrTran(aDados[nI, 11], '.', ''), ',', '.'))
        Z69->Z69_VALICM :=  Val(StrTran(StrTran(aDados[nI, 12], '.', ''), ',', '.'))
        Z69->Z69_VALIPI :=  Val(StrTran(StrTran(aDados[nI, 13], '.', ''), ',', '.'))
        Z69->Z69_VALPIS :=  Val(StrTran(StrTran(aDados[nI, 14], '.', ''), ',', '.'))
        Z69->Z69_VALCOF :=  Val(StrTran(StrTran(aDados[nI, 15], '.', ''), ',', '.'))
        Z69->Z69_MSGXML :=  AllTrim(aDados[nI, 19])     // IIF(!Empty(aDados[nI, 21]), AllTrim(aDados[nI, 19]) + ' - ' + AllTrim(aDados[nI, 21]), AllTrim(aDados[nI, 22]) + ' - ' + AllTrim(aDados[nI, 24])) 
        Z69->Z69_CAPTUR :=  IIF(Upper(AllTrim(aDados[nI, 21])) == 'SIM', 'S', 'N' )

        MsUnlock()

        nCount++

        Z69->( DBCloseArea() )
    
    Else		
        cLogAux += 'li ' + cValToChar( nI ) + ' - A chave está incorreta. Verifique se a chave está na quarta coluna e possui 44 caracteres.' + Chr(13) + Chr(10)
    EndIf

    IncProc()		
Next nI

SA1->( DBCloseArea() )
SA2->( DBCloseArea() )

MsgInfo( cValToChar( nCount ) + ' registros importados.' ) 

Return
