#INCLUDE 'PROTHEUS.CH'
#INCLUDE 'FWMVCDEF.CH'

User Function M460FIM()

Local aArea       := GetArea()
Local aPWiz       := {}
Local aRetWiz     := {}
Local cTotNota    := ""
Local cPerc       := ""
Local cVal        := Alltrim(Str(SE1->E1_VALOR))
Local cParc       := ""

If SC5->C5_XVLRTOT > 0
	                                             	
	cTotNota := SC5->C5_XVLRTOT - (SF2->F2_VALMERC + SF2->F2_DESCZFR) 
	cParc := U_xConPar(SE1->E1_PARCELA)
	
	MsgRun("Gerando títulos a receber","Aguarde...",{|| U_xProcRec(cTotNota, cParc/*, cPerc*/)})
	
EndIf

RestArea( aArea )

Return

//-----------------------------------------------------------------------------------------------------------

User Function xProcRec(cTotNota, cParc/*, cPerc*/)

Local nTotNota := cTotNota//Val(cTotNota)
Local nPerc    := 0//Val(cPerc)
Local aVetSE1  := {}
Local cNum     := " "
Local nVlr     := 0
Local nParc    := 0
Local nCont    := 0
Local nX       := 0
Local nValPa   := 0

Private lMsErroAuto := .F.

nParc := Val(cParc)

If nParc > 0
	nCont := nParc
Else
	nCont := 1
EndIf

If nCont == 1

	cNum := "9" + SE1->E1_NUM
	cNum := SubStr(cNum,1,10)
	cNum := Alltrim(cNum)
	
	nVlr := SE1->E1_VALOR
	//nVlr := Round(nPerc * (nVlr / 100),2)
	
	aVetSE1 := {}
	aAdd(aVetSE1, {"E1_FILIAL"  ,  FWxFilial("SE1"),  Nil})
	aAdd(aVetSE1, {"E1_NUM"     ,  SE1->E1_NUM     ,  Nil})
	aAdd(aVetSE1, {"E1_PREFIXO" ,  'RES'           ,  Nil})
	aAdd(aVetSE1, {"E1_PARCELA" ,  SE1->E1_PARCELA ,  Nil})
	aAdd(aVetSE1, {"E1_TIPO"    ,  SE1->E1_TIPO    ,  Nil})
	aAdd(aVetSE1, {"E1_NATUREZ" ,  SE1->E1_NATUREZ ,  Nil})
	aAdd(aVetSE1, {"E1_CLIENTE" ,  SE1->E1_CLIENTE ,  Nil})
	aAdd(aVetSE1, {"E1_LOJA"    ,  SE1->E1_LOJA    ,  Nil})
	aAdd(aVetSE1, {"E1_NOMCLI"  ,  SE1->E1_NOMCLI  ,  Nil})
	aAdd(aVetSE1, {"E1_EMISSAO" ,  SE1->E1_EMISSAO ,  Nil})
	aAdd(aVetSE1, {"E1_VENCTO"  ,  SE1->E1_VENCTO  ,  Nil})
	aAdd(aVetSE1, {"E1_VENCREA" ,  SE1->E1_VENCREA ,  Nil})
	aAdd(aVetSE1, {"E1_VALOR"   ,  nTotNota        ,  Nil})
	aAdd(aVetSE1, {"E1_VALJUR"  ,  SE1->E1_VALJUR  ,  Nil})
	aAdd(aVetSE1, {"E1_PORCJUR" ,  SE1->E1_PORCJUR ,  Nil})
	aAdd(aVetSE1, {"E1_ORIGEM"  ,  'M460FIM'       ,  Nil}) 
	aAdd(aVetSE1, {"E1_HIST"    ,  SE1->E1_HIST    ,  Nil})
	aAdd(aVetSE1, {"E1_MOEDA"   ,   1              ,  Nil})
	aAdd(aVetSE1, {"E1_VEND1"   ,  SE1->E1_VEND1   ,  Nil})
	aAdd(aVetSE1, {"E1_VEND2"   ,  SE1->E1_VEND2   ,  Nil})
	aAdd(aVetSE1, {"E1_VEND3"   ,  SE1->E1_VEND3   ,  Nil})
	aAdd(aVetSE1, {"E1_VEND4"   ,  SE1->E1_VEND4   ,  Nil})
	aAdd(aVetSE1, {"E1_VEND5"   ,  SE1->E1_VEND5   ,  Nil})    ¦
	aAdd(aVetSE1, {"E1_COMIS1"  ,  SE1->E1_COMIS1  ,  Nil})
	aAdd(aVetSE1, {"E1_COMIS2"  ,  SE1->E1_COMIS2  ,  Nil})
	aAdd(aVetSE1, {"E1_COMIS3"  ,  SE1->E1_COMIS3  ,  Nil})
	aAdd(aVetSE1, {"E1_COMIS4"  ,  SE1->E1_COMIS4  ,  Nil})
	aAdd(aVetSE1, {"E1_COMIS5"  ,  SE1->E1_COMIS5  ,  Nil})
	
	
	//Begin Transaction
 		lMsErroAuto := .F.
		MSExecAuto({|x,y| FINA040(x,y)}, aVetSE1, 3)
		   
	    lMsErroAuto := .F.
	     
	    If lMsErroAuto
	        MostraErro()
	        DisarmTransaction()
	    Else
	    	MsgInfo("Título a receber incluido no valor de: R$" + Alltrim(Str(nTotNota)) + "","Título a Receber!")
	    EndIf

	//End Transaction
	
Else
	nValPa := nTotNota / nParc
	
	For nX := 1 To nParc
	
		cNum :=  "9" + Alltrim(Str(nX)) + SE1->E1_NUM
		cNum := SubStr(cNum,1,9)
		cNum := Alltrim(cNum)
		
		nVlr := SE1->E1_VALOR 
		//nVlr := Round(nPerc * (nVlr / 100),2)
		
		aVetSE1 := {}
		
		aAdd(aVetSE1, {"E1_FILIAL"  ,  FWxFilial("SE1"),  Nil})
		aAdd(aVetSE1, {"E1_NUM"     ,  SE1->E1_NUM     ,  Nil})
		aAdd(aVetSE1, {"E1_PREFIXO" ,  'RE' + Alltrim(Str(nX)),  Nil})
		aAdd(aVetSE1, {"E1_PARCELA" ,  U_xRetPar(nX)   ,  Nil})
		aAdd(aVetSE1, {"E1_TIPO"    ,  SE1->E1_TIPO    ,  Nil})
		aAdd(aVetSE1, {"E1_NATUREZ" ,  SE1->E1_NATUREZ ,  Nil})
		aAdd(aVetSE1, {"E1_CLIENTE" ,  SE1->E1_CLIENTE ,  Nil})
		aAdd(aVetSE1, {"E1_LOJA"    ,  SE1->E1_LOJA    ,  Nil})
		aAdd(aVetSE1, {"E1_NOMCLI"  ,  SE1->E1_NOMCLI  ,  Nil})
		aAdd(aVetSE1, {"E1_EMISSAO" ,  SE1->E1_EMISSAO ,  Nil})
		aAdd(aVetSE1, {"E1_VENCTO"  ,  U_xVencto(SE1->E1_FILIAL,SE1->E1_NUM,nX)  ,  Nil})
		aAdd(aVetSE1, {"E1_VENCREA" ,  U_xVencre(SE1->E1_FILIAL,SE1->E1_NUM,nX) ,  Nil})
		aAdd(aVetSE1, {"E1_VALOR"   ,  nValPa          ,  Nil})
		aAdd(aVetSE1, {"E1_VALJUR"  ,  SE1->E1_VALJUR  ,  Nil})
		aAdd(aVetSE1, {"E1_PORCJUR" ,  SE1->E1_PORCJUR ,  Nil})
		aAdd(aVetSE1, {"E1_ORIGEM"  ,  'M460FIM'       ,  Nil}) 
		aAdd(aVetSE1, {"E1_HIST"    ,  SE1->E1_HIST    ,  Nil})
		aAdd(aVetSE1, {"E1_MOEDA"   ,   1              ,  Nil})
		aAdd(aVetSE1, {"E1_VEND1"   ,  SE1->E1_VEND1   ,  Nil})
		aAdd(aVetSE1, {"E1_VEND2"   ,  SE1->E1_VEND2   ,  Nil})
		aAdd(aVetSE1, {"E1_VEND3"   ,  SE1->E1_VEND3   ,  Nil})
		aAdd(aVetSE1, {"E1_VEND4"   ,  SE1->E1_VEND4   ,  Nil})
		aAdd(aVetSE1, {"E1_VEND5"   ,  SE1->E1_VEND5   ,  Nil})
		aAdd(aVetSE1, {"E1_COMIS1"  ,  SE1->E1_COMIS1  ,  Nil})
		aAdd(aVetSE1, {"E1_COMIS2"  ,  SE1->E1_COMIS2  ,  Nil})
		aAdd(aVetSE1, {"E1_COMIS3"  ,  SE1->E1_COMIS3  ,  Nil})
		aAdd(aVetSE1, {"E1_COMIS4"  ,  SE1->E1_COMIS4  ,  Nil})
		aAdd(aVetSE1, {"E1_COMIS5"  ,  SE1->E1_COMIS5  ,  Nil})
		
		//Begin Transaction

		    lMsErroAuto := .F.
		    MSExecAuto({|x,y| FINA040(x,y)}, aVetSE1, 3)
		     
		    If lMsErroAuto
		        MostraErro()
		        DisarmTransaction()
		    Else
		    	
		    	MsgInfo("Título a receber incluido no valor de: R$" + Alltrim(Str(nValPa)) + "","Título a Receber!")
		    EndIf

		//End Transaction
		
	Next nX
	
EndIf
 

Return


User Function xConPar(cParc)
	
	Local _Letras:="ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	Local _nParc := 0
	
	cParc := Alltrim(cParc)
	
	_nParc := At(cParc,_Letras)
	
	_nParc := Str(_nParc,1)

Return(_nParc)

User Function xRetPar(nParc)
	
	Local cParc   := ""
	
	If nParc == 1
		cParc := "A"
	ElseIf nParc == 2
		cParc := "B"
	ElseIf nParc == 3
		cParc := "C"
	ElseIf nParc == 4
		cParc := "D"
	ElseIf nParc == 5
		cParc := "E"
	ElseIf nParc == 6
		cParc := "F"
	ElseIf nParc == 7
		cParc := "G"
	ElseIf nParc == 8
		cParc := "H"
	ElseIf nParc == 9
		cParc := "I"
	ElseIf nParc == 10
		cParc := "J"
	ElseIf nParc == 11
		cParc := "K"
	ElseIf nParc == 12
		cParc := "L"
	ElseIf nParc == 13
		cParc := "M"
	ElseIf nParc == 14
		cParc := "N"
	ElseIf nParc == 15
		cParc := "O"
	ElseIf nParc == 16
		cParc := "P"
	ElseIf nParc == 17
		cParc := "Q"
	ElseIf nParc == 18
		cParc := "R"
	ElseIf nParc == 19
		cParc := "S"
	ElseIf nParc == 20
		cParc := "T"
	ElseIf nParc == 21
		cParc := "U"
	ElseIf nParc == 22
		cParc := "V"
	ElseIf nParc == 23
		cParc := "W"
	ElseIf nParc == 24
		cParc := "X"
	ElseIf nParc == 25
		cParc := "Y"
	ElseIf nParc == 26
		cParc := "Z"
	EndIf
	

Return cParc


User Function xVencto(cFil,cNum,nParcela)
	
	Local aArea    := GetArea()
	Local cParcela := ""
	Local dData    
	
	cParcela := U_xRetPar(nParcela)
	
	DbSelectArea("SE1")
	DbSetOrder(1)
	
	If DbSeek(cFil + '1  ' + cNum + cParcela)
	
		dData := SE1->E1_VENCTO
	
	EndIf
	
	RestArea( aArea )

Return dData


User Function xVencre(cFil,cNum,nParcela)
	
	Local aArea    := GetArea()
	Local cParcela := ""
	Local dData    
	
	cParcela := U_xRetPar(nParcela)
	
	DbSelectArea("SE1")
	DbSetOrder(1)
	
	If DbSeek(cFil + '1  ' + cNum + cParcela)
	
		dData := SE1->E1_VENCREA
	
	EndIf
	
	RestArea( aArea )

Return dData