#INCLUDE "TOTVS.CH"
#INCLUDE "PROTHEUS.CH"

User Function AtuPrc()     

Local cQuery     := ""
Local cAliAux    := GetNextAlias()
Local nRecno     := 0

cQuery := "SELECT R_E_C_N_O_, * FROM DA1010 WHERE DA1_CODTAB = 'E09' AND D_E_L_E_T_ = ' '"

cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

(cAliAux)->(dbGoTop())

Do While (cAliAux)->(!Eof())

	nRecno := (cAliAux)->R_E_C_N_O_
	
	DA1->(DbGoto(nRecno))
	
	RecLock("DA1",.F.)
	
		DA1->DA1_PRCVEN := 30 * (DA1->DA1_PRCVEN/100)
		
	MsUnLock()
	
	(cAliAux)->(DbSkip())
EndDo

Alert ("Concluido")

Return