#include 'protheus.ch'
#include 'rwmake.ch'
#include 'fivewin.ch'
#include 'tbiconn.ch'
#include 'fileio.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: SLCF1080      || Autor: Lucas Rocha          || Data: 15/03/18  ||
||-------------------------------------------------------------------------||
|| Descrição: Workflow de pedidos de compra já aprovados 		   ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function SLCF1080()

If IsBlind()
	PREPARE ENVIRONMENT EMPRESA "01" FILIAL "01" MODULO "COM"  USER "workflow" PASSWORD "workflow"
EndIf

Private cServer		:= GetMV( 'MV_RELSERV' )
Private cUser		:= GetMV( 'MV_RELACNT' )
Private cPass		:= GetMV( 'MV_RELPSW' )
Private lAuth		:= GetMV( 'MV_RELAUTH' )
Private cGrpMail1	:= SuperGetMv( 'ES_WFPCA1', .T., 'lucas.rocha,' )	// Pessoas do grupo da controladoria. 	Ex.: alice.goebel
Private cGrpMail2	:= SuperGetMv( 'ES_WFPCA2', .T., 'lucas.rocha,' )	// Pessoas do grupo de compras. 	Ex.: lauro.fonseca
Private cSubject	:= 'Pedidos Aprovados'
Private cBCC		:= '' 
Private cTexto		:= ''
Private aFile 		:= {}
Private aDados 		:= {}
Private cDir 		:= '\controle_wf\'
Private cArq 		:= 'pc_aprov.txt'
Private nHandle		:= 0
Private count1 		:= 2
Private count2 		:= 2
Private cHtml, cHtmlM1, cHtmlM1F, cHtmlM2, cHtmlM2F, cHtmlT1, cHtmlT2, cHtmlF, cCompr, lAux1, lAux2
                                 
aFile 	:= Directory( cDir + cArq, 'D' ) 
cHtmlT1 := ''
cHtmlT2 := ''

// Começa verificando a pasta e o arquivo de controle dos pedidos que já foram enviados por e-mail para não repetí-los

If !ExistDir( cDir )			// Verifica se não existe a pasta criada
	
	If !CriaDir() .or. !CriaArq()	// Cria a pasta e o arquivo
		Return
	EndIf

ElseIf  !File( cDir + cArq )		// Verifica se não existe o arquivo criado

	If !CriaArq()			// Cria o arquivo
		Return
	EndIf
		
ElseIf aFile[1,3] <= DATE() - 7 	// Verifica se já faz uma semana desde que o arquivo foi criado	    
    
    If !DeletaArq() .or. !CriaArq()	// Deleta o arquivo velho e recria um novo
    	Return
    EndIf 
    			      			
Else                                	// O arquivo está ok
				
	If !LeiaArq()   		// Captura os pedidos que já foram enviados e alimenta o array aDados
		Return
	EndIf			

EndIf	


// Começa a montar o código HTML ...
cHtml := " <!DOCTYPE html> "
cHtml += " <html lang='pt-br'> "
cHtml += " 	<head> "
cHtml += " 		<title>Pedidos de Compra Aprovados</title> "
cHtml += " 		<meta charset='utf-8'> "
cHtml += " 		<style type='text/css'> "
cHtml += " 			h3 { "
cHtml += " 				font-family: calibri; "
cHtml += " 			} "
cHtml += " 			table { "
cHtml += " 				border-collapse: collapse; "
cHtml += " 				width: 100%; "
cHtml += " 				max-width: 1400px; "
cHtml += " 				min-width: 900px; "
cHtml += " 				font-family: calibri; "
cHtml += " 			} "
cHtml += " 			th { "
cHtml += " 				font-size: 16px; "
cHtml += " 				color: white; " 
cHtml += " 				background-color: #ff3333; "
cHtml += " 				font-weight: bold; " 
cHtml += " 				padding: 6px; "
cHtml += " 			} "
cHtml += " 			td { "
cHtml += " 				font-size: 12px; " 
cHtml += " 				padding: 8px; "
cHtml += " 			} "
cHtml += " 			th, td { "
cHtml += " 				border-bottom: 1px solid #ddd; "
cHtml += " 				text-align: center; "
cHtml += " 			} "
cHtml += " 		</style> "
cHtml += " 	</head> "
cHtml += " 	<body> "

// Cabeçalho 1
cHtmlM1 := "		<div> "
cHtmlM1 += "			<h3>Relação de pedidos liberados do dia " + DTOC( DATE() ) + ".</h3> "
cHtmlM1 += "			<br> "
cHtmlM1 += "		</div> "
cHtmlM1 += " 		<table> "
cHtmlM1 += " 			<thead> "
cHtmlM1 += " 				<tr bgcolor='#ff0000c2''> "
cHtmlM1 += " 					<th>Filial</th> "
cHtmlM1 += " 					<th>Nº Pedido</th> "
cHtmlM1 += " 					<th>Data Digit.</th> "
cHtmlM1 += " 					<th>Fornecedor</th> "
cHtmlM1 += " 					<th>Total</th> "  
cHtmlM1 += " 					<th>Comprador</th> "
cHtmlM1 += " 				</tr> "
cHtmlM1 += " 			</thead> "
cHtmlM1 += " 			<tbody> "

cHtmlM1F := " 			</tbody> "
cHtmlM1F += " 		</table> "

// Cabeçalho 2
cHtmlM2  := "		<p>Não há pedidos novos para mostrar.</p>"
cHtmlM2F := ""

// Fim
cHtmlF := " 	</body> "
cHtmlF += " </html> " 


// Monta a query pra trazer todos os pedidos de compra que foram aprovados e ainda não tenham sido enviados por e-mail
cQuery := "  SELECT C7_FILIAL, C7_NUM, C7_EMISSAO, A2_NOME, CR_TOTAL, C7_USER, C7_FORNECE, C7_LOJA, MAX(CR_DATALIB) AS DATA, C7_CHVZ31  "
cQuery += "  FROM SC7010  "
cQuery += "  	JOIN SA2010  "
cQuery += "  		ON A2_COD = C7_FORNECE  "
cQuery += "  		AND A2_LOJA = C7_LOJA  "
cQuery += "  		AND SA2010.D_E_L_E_T_ = ''  "
cQuery += " 	JOIN SCR010 "
cQuery += " 		ON CR_FILIAL = C7_FILIAL "
cQuery += " 		AND CR_NUM = C7_NUM "
cQuery += " 		AND CR_DATALIB <> '' "
cQuery += " 		AND SCR010.D_E_L_E_T_ = '' "
cQuery += "  WHERE C7_CONAPRO = 'L'  "
cQuery += "  	AND SC7010.D_E_L_E_T_ = ''  "
cQuery += "  	AND C7_QUJE = 0  "
cQuery += "  	AND C7_RESIDUO = ''  "
cQuery += "  	AND CR_DATALIB >= '" + DTOS( DATE() - 7 ) + "' "
cQuery += "  	AND C7_ENCER <> 'E'  "

For i := 1 to len( aDados )		// Não traz os pedidos que estão no arquivo de controle, ou seja, os que já foram enviados por e-mail
	cQuery += " AND NOT ( C7_FILIAL = '" + aDados[i,1] + "' AND C7_NUM = '" + aDados[i,2] + "' AND C7_FORNECE = '" + aDados[i,3] + "' AND C7_LOJA = '" + aDados[i,4] + "' ) "
Next                                        

cQuery += " GROUP BY C7_FILIAL, C7_NUM, C7_EMISSAO, A2_NOME, CR_TOTAL, C7_USER, C7_FORNECE, C7_LOJA, C7_CHVZ31 "	
cQuery += " ORDER BY C7_FILIAL, C7_NUM "   


dbUseArea( .t., 'TOPCONN', TCGenQry(,, cQuery ), 'TMP', .t., .t. )		// Abre a conexão da query com os pedidos

TMP->( dbGoTop() )
                        
While TMP->( !EOF() )
    	cCompr := ' - '
    	lAux1  := .F.
    	lAux2  := .F.
    
    	// Transforma o código de usuário no nome de login e alimenta a variável cCompr
    	PswOrder( 1 )	
	If PswSeek( TMP->C7_USER, .T. )
		 
		cCompr := PswRet()[1][2]				
	EndIf                   
	
	// Recebe apenas os pedidos que algúem do grupo controladoria incluiu
    	If cCompr $ cGrpMail1	 
        
        	cHtmlT1  += IIf ( count1 % 2 == 1, "	<tr bgcolor='#f2f2f2'> ", "	<tr> " ) 	// Design difierenciado		
	    	cHtmlT1  += AlimentaTabela()
	    
	    	count1 	 := count1 + 1
	    	lAux1 := .T.
	    	     	       	
   	EndIf
   	
   	// Recebe apenas os pedidos que algúem do grupo compras incluiu ou os pedidos que foi preenchido o cód. de investidor
	If !Empty( TMP->C7_CHVZ31 ) .OR. cCompr $ cGrpMail2
	    
		cHtmlT2 += IIf ( count2 % 2 == 1, "	<tr bgcolor='#f2f2f2'> ", "	<tr> " ) 	// Design difierenciado
		cHtmlT2 += AlimentaTabela()
		
		count2 := count2 + 1
		lAux2 := .T.
		
	EndIf 
	
	If lAux1 .or. lAux2 		
		// Novos pedidos para o arquivo de controle			
		cTexto  +=  TMP->C7_FILIAL + "," + TMP->C7_NUM + "," + TMP->C7_FORNECE + "," + TMP->C7_LOJA + Chr(13) + Chr(10)				
	EndIf
	
	TMP->( dbSkip() )
End    

dbCloseArea('TMP') 
     

// Ínicio da separação dos HTMLs por grupo de e-mail

// 		Estrutura:  
cHtml1 := cHtml
cHtml2 := cHtml
// 		Cabeçalho Vazio/Tabela: 
cHtml1 += IIf( count1 == 2, cHtmlM2, cHtmlM1 )
cHtml2 += IIf( count2 == 2, cHtmlM2, cHtmlM1 )
// 		Itens Tabela: 
cHtml1 += IIf( Empty( cHtmlT1 ), '', cHtmlT1 )
cHtml2 += IIf( Empty( cHtmlT2 ), '', cHtmlT2 )
// 		Fechar HTML: 
cHtml1 += IIf( count1 == 2, cHtmlM2F, cHtmlM1F ) + cHtmlF
cHtml2 += IIf( count2 == 2, cHtmlM2F, cHtmlM1F ) + cHtmlF


// Monta os e-mails que receberão o workflow a partir dos parâmetros ES_WFPCA1 e ES_WFPCA2
cTo1 := Monta_cTo( cGrpMail1 )
cTo2 := Monta_cTo( cGrpMail2 )

If !EnviaEmail()
	Return
EndIf


// Início da gravçaão dos novos pedidos no arquivo controle
If !GravaArq() 
	Return
EndIf


If IsBlind()
	Reset Environment
EndIf


Return

*********************************************************************************************************
Static Function Monta_cTo ( cParam )

Local cTo := ""

If RIGHT( cParam , 1 ) != "," .AND. !Empty( cParam )
	
	cParam += ","
	
EndIf

If "," $ cParam
	
	cTo := StrTran( cParam, ",", "@slcalimentos.com.br; " ) 
	
EndIf

Return cTo

*********************************************************************************************************
Static Function CriaArq()

nHandle := FCreate( cDir + cArq )	// Cria o arquivo
	
If nHandle < 0
	// MsgAlert( 'Erro durante a recriação do arquivo!' )
	Return .F.
EndIf                       

FClose( nHandle )

Return .T.

*********************************************************************************************************
Static Function CriaDir()
	
If MakeDir( cDir ) != 0 		// Cria a pasta
	// MsgAlert( 'Erro durante a criação da pasta!' )
	Return .F.        
EndIf

If !CriaArq()
	Return .F.
EndIf
	
Return .T.

*********************************************************************************************************
Static Function DeletaArq()

If FErase(cDir + cArq) < 0 	  	// Deleta o arquivo
	// MsgStop( 'Erro durante a deleção do arquivo!' )
	Return .F.
Endif

Return .T.

*********************************************************************************************************
Static Function LeiaArq()

nHandle := FOpen( cDir + cArq )

If nHandle < 0
	// MsgAlert( 'Erro durante a abertura do arquivo!' )
	Return .F.
Endif

FT_FUse( cDir + cArq )
FT_FGoTop()
 
While !FT_FEOF()
              
	If !Empty( FT_FREADLN() )
		aadd( aDados, Separa( FT_FREADLN(), ",", .T. ) )   // Separa dentro do array aDados: Filial [1] - Número [2] - Fornecedor [3] - Loja [4]
		FT_FSKIP()
	EndIf
	
EndDo
 
FT_FUse()
FClose( nHandle )

Return .T.

*********************************************************************************************************
Static Function GravaArq()

nHandle := FOpen( cDir + cArq, FO_READWRITE + FO_SHARED )	// Abre o arquivo para editá-lo

If nHandle < 0
	// MsgAlert( 'Erro de abertura : FERROR ' + str( FError(), 4 ) )
	Return .F.
EndIf
 
FSeek(nHandle, 0, FS_END) 
FWrite( nHandle, cTexto )
FClose( nHandle ) 


Return .T.

*********************************************************************************************************
Static Function AlimentaTabela()

Local cHtmlTmp := ''

cHtmlTmp += " 		<td>" + TMP->C7_FILIAL + "</td> " 
cHtmlTmp += " 	   	<td>" + TMP->C7_NUM + "</td> " 
cHtmlTmp += " 	   	<td>" + DTOC( STOD( TMP->C7_EMISSAO ) ) + "</td> " 
cHtmlTmp += " 	   	<td>" + TMP->A2_NOME + "</td> "
cHtmlTmp += " 	   	<td>R$ " + Transform(TMP->CR_TOTAL, "@E 99,999,999,999.99") + "</td> "
cHtmlTmp += " 	   	<td>" + cCompr + "</td> "
cHtmlTmp += " 	</tr> "

Return cHtmlTmp

*********************************************************************************************************
Static Function EnviaEmail()

// Conectando-se com o servidor de e-mail
CONNECT SMTP SERVER cServer ACCOUNT cUser PASSWORD cPass RESULT lResult

// Autenticando
If lResult .And. lAuth
	lResult := MailAuth( cUser, cPass )
	If !lResult
		lResult := QADGetMail() // funcao que abre uma janela perguntando o usuario e senha para fazer autenticacao
	EndIf
EndIf

If !lResult
	GET MAIL ERROR cError
	MsgAlert( 'Erro de Autenticacao no Envio de e-mail antes do envio: ' + cError, cSubject )
	Return .F.
EndIf

If cTo1 != ''
	SEND MAIL FROM 'workflowal@slcalimentos.com.br' TO cTo1 BCC cBCC SUBJECT cSubject BODY cHtml1 /*FORMAT TEXT*/ RESULT lResult
EndIf
If cTo2 != ''
	SEND MAIL FROM 'workflowal@slcalimentos.com.br' TO cTo2 BCC cBCC SUBJECT cSubject BODY cHtml2 /*FORMAT TEXT*/ RESULT lResult
EndIf

If !lResult
	GET MAIL ERROR cError
	MsgAlert( 'Erro de Autenticacao no Envio de e-mail depois do envio: ' + cError, cSubject )
	Return .F.
EndIf

DISCONNECT SMTP SERVER

Return .T.
