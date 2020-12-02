#INCLUDE "TOTVS.CH"
#INCLUDE "PROTHEUS.CH"

User Function OM010MNU()

AAdd(aRotina, {"Aumento Anual", "U_xAumAnu", 0, 3})
//AAdd(aRotina, {"Aumento Individual", "U_xAumInd", 0, 3})

Return 


User Function xAumAnu()

Local aPWiz   := {}
Local aRetWiz := {}
Local cPerc   := ""
Local nVal    := 0
Local aArea   := GetArea()

aAdd(aPWiz,{ 1,"Percentual de aumento:"    ,Space(3) ,"","","","", ,.T.})

aAdd(aRetWiz,Space(3))

ParamBox(aPWiz,"Aumento Anual - Elitte",@aRetWiz,,,,,,) 

cPerc   := Alltrim(aRetWiz[1])

If Empty(cPerc)
	MsgInfo("Percentual não informado","Cód. Invalido")
	Return
EndIf

nVal := Val(cPerc)

MsgRun("Atualizando preços...","Aguarde...",{|| U_xAumento(nVal)})

RestArea( aArea )

Return 

User Function xAumento(nVal)

Local aArea   := GetArea()
Local nPrc    := 0

DbSelectArea("DA1")
DbSetOrder(1)
DbGoTop()

While DA1->(!Eof())

	RecLock("DA1",.F.)
	
		nPrc := nVal * (DA1->DA1_PRCVEN/100)
		
		DA1->DA1_PRCVEN += nPrc
		
	MsUnLock()
	
	DA1->(DbSkip())
	
EndDo

MsgInfo("Atualização Efetuada","Atualização")

RestArea( aArea )

Return 

User Function xAumInd()

Local aPWiz   := {}
Local aRetWiz := {}
Local cPerc   := ""
Local nVal    := 0
Local aArea   := GetArea()

aAdd(aPWiz,{ 1,"Percentual de aumento:"    ,Space(3) ,"","","","", ,.T.})

aAdd(aRetWiz,Space(3))

ParamBox(aPWiz,"Aumento Individual - Elitte",@aRetWiz,,,,,,) 

cPerc   := Alltrim(aRetWiz[1])

If Empty(cPerc)
	MsgInfo("Percentual não informado","Cód. Invalido")
	Return
EndIf

nVal := Val(cPerc)

MsgRun("Atualizando preços...","Aguarde...",{|| U_xIndAum(nVal)})

RestArea( aArea )

Return 

User Function xIndAum(nVal)

Local aArea   := GetArea()

MsgInfo("Atualização Efetuada","Atualização")

RestArea( aArea )

Return