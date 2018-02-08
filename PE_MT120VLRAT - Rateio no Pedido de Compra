#Include 'Protheus.ch'

/*___________________________________________________________________________
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
||-------------------------------------------------------------------------||
|| Função: MT120VLRAT     || Autor: Lucas Rocha         || Data: 13/12/17  ||
||-------------------------------------------------------------------------||
|| Descrição: PE que valida a linha do rateio por CC no pedido de compras  ||
||-------------------------------------------------------------------------||
|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯*/

User Function MT120VLRAT()

Local lRet			:=	.T.
Local aColsSCH		:=	ParamIXB[1]
Local aHeaderSCH	:=	ParamIXB[2]
Local nLinha		:=	ParamIXB[3]

If Empty(aColsSCH[n][3]) .OR. Empty(aColsSCH[n][4]) .OR. Empty(aColsSCH[n][5]) 
	MsgAlert( 'O Centro de Custo, a Conta Contábil e o Item Contábil não podem ficar em branco.' )
	
	lRet := .F. 
EndIf

Return (lRet)
