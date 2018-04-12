#include 'protheus.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: MT103SE2      || Autor: Lucas Rocha         || Data: 10/04/18   ||
||-------------------------------------------------------------------------||
|| Descrição: Possibilita a adição de campos ao aCols da aba duplicatas no ||
|| Documento de Entrada.   						   						   ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function MT103SE2()

//Local aHead:= PARAMIXB[1]			// Vetor contendo os registros adicionados por padrão para o aHeader de título financeiro
//Local lVisual:= PARAMIXB[2]		// Variável lógica que determina se a operação é de visualização (.T.) ou não (.F.).
Local aRet := {}	   				// Customizações desejadas para adição do campo no grid de informações

If MsSeek( 'E2_VENCREA' )			// Vencimento original
   
	AADD( aRet, { ;
			TRIM( X3Titulo() ), SX3->X3_CAMPO, SX3->X3_PICTURE, SX3->X3_TAMANHO, ;
			SX3->X3_DECIMAL, '', SX3->X3_USADO, SX3->X3_TIPO, SX3->X3_F3, SX3->X3_CONTEXT, ;
			SX3->X3_CBOX, SX3->X3_RELACAO, ;
			'.T.' }) 
	
EndIf 


Return aRet
