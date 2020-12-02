#INCLUDE "TOTVS.CH"
#INCLUDE "PROTHEUS.CH"

User Function xRecTit()     

Local aPWiz     := {}
Local aRetWiz   := {}
Local cFilDe    := ""
Local cFilAte   := "" 
Local cConDe    := "" 
Local dDtDe     := CTOD("//")
Local dDtAte    := CTOD("//")
Local cConci    := 0
Local cHist     := ""
           
Private lInverte  := .F. 
Private cMark     := GetMark()    
Private oMark 
Private oProcess
Private nCont     := 0
Private lVinOk    := .F.
Private cVlrSal   := ""
Private oProcExc

aAdd(aPWiz,{ 1,"Filial de: "             ,Space(TamSX3("CT2_FILIAL")[1]),"","","SM0","",9   ,.T.})
aAdd(aPWiz,{ 1,"Filial ate: "            ,Space(TamSX3("CT2_FILIAL")[1]) ,"","","SM0","",9   ,.T.})
aAdd(aPWiz,{ 1,"Data de: "               ,Ctod("") ,"","",""   ,  ,60 ,.T.})
aAdd(aPWiz,{ 1,"Data ate: "              ,Ctod("") ,"","",""   ,  ,60,.T.})

aAdd(aRetWiz,Space(TamSX3("E1_FILIAL")[1]))
aAdd(aRetWiz,Space(TamSX3("E1_FILIAL")[1]))
aAdd(aRetWiz,Ctod(""))
aAdd(aRetWiz,Ctod(""))

ParamBox(aPWiz,"Elitte",@aRetWiz,,,,,,) 

cFilDe  := Alltrim(aRetWiz[1])
cFilAte := Alltrim(aRetWiz[2]) 
dDtDe   := aRetWiz[3] 
dDtAte  := aRetWiz[4] 

oProcess := MsNewProcess():New( { || XProcRel(cFilde, cFilAte, dDtDe, dDtAte) } , "Carregando tabela temporária" , "Aguarde..." , .F. )
oProcess:Activate()

//-------------------------------------------------------------------

Function XProcRel(cFilde, cFilAte, dDtDe, dDtAte)

Local cArquivo  := "CONCI.XLS"
Local oExcelApp := Nil
Local cPath     := "C:\Percentual"
Local nTotal    := 0
Local oExcel
Local oExcelApp
Local _oPlan
Local cQuery    := ""
Local cAliAux   := GetNextAlias()
Local aCampos   := {}
Local cArqTrb   := GetNextAlias()

aColunas := {}
aLocais  := {} 
oBrush1  := TBrush():New(, RGB(193,205,205))

AADD(aCampos,{"E1_FILIAL"   ,"C",TamSX3("E1_FILIAL")[1],0})
AADD(aCampos,{"E1_PREFIXO"  ,"C",TamSX3("E1_PREFIXO"  )[1],0})
AADD(aCampos,{"E1_NUM"      ,"C",TamSX3("E1_NUM"  )[1],0})
AADD(aCampos,{"E1_PARCELA"  ,"C",TamSX3("E1_PARCELA")[1],0})
AADD(aCampos,{"E1_TIPO"     ,"C",TamSX3("E1_TIPO"   )[1],0})
AADD(aCampos,{"E1_NATUREZ"  ,"C",TamSX3("E1_NATUREZ")[1],0})
AADD(aCampos,{"E1_CLIENTE"  ,"C",TamSX3("E1_CLIENTE")[1],0})
AADD(aCampos,{"E1_LOJA"     ,"C",TamSX3("E1_LOJA"    )[1],0})
AADD(aCampos,{"E1_NOMCLI"   ,"C",TamSX3("E1_NOMCLI")[1],0})
AADD(aCampos,{"E1_EMISSAO"  ,"D",TamSX3("E1_EMISSAO")[1],0})
AADD(aCampos,{"E1_VENCTO"   ,"D",TamSX3("E1_VENCTO" )[1],2})
AADD(aCampos,{"E1_VENCREA"  ,"D",TamSX3("E1_VENCREA"  )[1],0})
AADD(aCampos,{"E1_VALOR"    ,"N",TamSX3("E1_VALOR")[1],2})
AADD(aCampos,{"E1_HIST"     ,"C",TamSX3("E1_HIST")[1],0})

If !ApOleClient('MsExcel')

    MsgAlert("Falha ao abrir Excel!")
    Return

EndIf

oExcel  := FWMSExcel():New()
cAba    := "Títulos a Receber"
cTabela := "Títulos a Receber - ELITTE"

// Criação de nova aba 
oExcel:AddworkSheet(cAba)

// Criação de tabela
oExcel:AddTable (cAba,cTabela)

// Criação de colunas 
oExcel:AddColumn(cAba,cTabela,"FILIAL"        	 ,1,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"PREFIXO"  		 ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"NUMERO"           ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"PARCELA"     	 ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"TIPO"   	         ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"NATUREZA"         ,1,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"CLIENTE"          ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"LOJA"    		 ,2,1,.F.) 
oExcel:AddColumn(cAba,cTabela,"NOME CLIENTE"     ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"DATA EMISSAO"     ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"DATA VENCTO."     ,3,2,.F.) 
oExcel:AddColumn(cAba,cTabela,"DATA VENCREA."  	 ,3,2,.F.)
oExcel:AddColumn(cAba,cTabela,"VALOR"  		     ,2,1,.F.)
oExcel:AddColumn(cAba,cTabela,"HISTORICO"  		 ,3,2,.F.)


cQuery := "SELECT * FROM "
cQuery += RetSqlName("SE1") + " SE1 "
cQuery += " WHERE "
cQuery += " E1_FILIAL   Between '" + cFilDe    + "' AND '" + cFilAte  + "' " 
cQuery += " AND E1_EMISSAO  >= '" + dtos(dDtDe)  + "' AND E1_EMISSAO <= '" + dtos(dDtAte) + "' " 
cQuery += " AND E1_ORIGEM = 'M460FIM'"
cQuery += " AND D_E_L_E_T_ = ' '"

cQuery := ChangeQuery(cQuery) 
 
dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cAliAux,.T.,.T.)

//oProcExc:SetRegua1( (cAliAux)->(RecCount()) ) //Alimenta a primeira barra de progresso
//Impressão Relatório

If _oPlan <> Nil
	_oPlan:Delete() 
	_oPlan := Nil
EndIf

// Criando o objeto do arquivo temporário
_oPlan := FwTemporaryTable():New(cArqTrb)

// Criando a estrutura do objeto  
_oPlan:SetFields(aCampos)

// Criando o indice da tabela
_oPlan:AddIndex("1",{"E1_PREFIXO"})

_oPlan:Create()


Do While !(cAliAux)->(Eof())
		//oProcExc:IncRegua1("Processando tabela temporária")	
		RecLock(cArqTrb,.T.)
		(cArqTrb)->E1_FILIAL    := Alltrim((cAliAux)->E1_FILIAL)
		(cArqTrb)->E1_PREFIXO   := Alltrim((cAliAux)->E1_PREFIXO) 
		(cArqTrb)->E1_NUM       := Alltrim((cAliAux)->E1_NUM)
		(cArqTrb)->E1_PARCELA   := Alltrim((cAliAux)->E1_PARCELA)
		(cArqTrb)->E1_TIPO      := Alltrim((cAliAux)->E1_TIPO)
		(cArqTrb)->E1_NATUREZ   := Alltrim((cAliAux)->E1_NATUREZ)
		(cArqTrb)->E1_CLIENTE   := Alltrim((cAliAux)->E1_CLIENTE)
		(cArqTrb)->E1_LOJA      := Alltrim((cAliAux)->E1_LOJA)
		(cArqTrb)->E1_NOMCLI    := Alltrim((cAliAux)->E1_NOMCLI)
		(cArqTrb)->E1_EMISSAO   := stod((cAliAux)->E1_EMISSAO)
		(cArqTrb)->E1_VENCTO    := stod((cAliAux)->E1_VENCTO)
		(cArqTrb)->E1_VENCREA   := stod((cAliAux)->E1_VENCREA)
		(cArqTrb)->E1_VALOR     := (cAliAux)->E1_VALOR
		(cArqTrb)->E1_HIST      := Alltrim((cAliAux)->E1_HIST)
		MsUnLock()
		(cAliAux)->(dbskip())
	Enddo
	
DbselectArea(cArqTrb)
dbGoTop()

While !(cArqTrb)->(Eof())
 	//oProcExc:IncRegua2("Gerando a planilha...")	
    // Criação de Linhas 
    oExcel:AddRow(cAba,cTabela, { (cArqTrb)->E1_FILIAL   ,;
                                  (cArqTrb)->E1_PREFIXO  ,; 
                                  (cArqTrb)->E1_NUM      ,; 
                                  (cArqTrb)->E1_PARCELA  ,;
                                  (cArqTrb)->E1_TIPO     ,;
                                  (cArqTrb)->E1_NATUREZ  ,;
                                  (cArqTrb)->E1_CLIENTE  ,;
                                  (cArqTrb)->E1_LOJA     ,;
                                  (cArqTrb)->E1_NOMCLI   ,;
                                  (cArqTrb)->E1_EMISSAO  ,;
                                  (cArqTrb)->E1_VENCTO   ,;
                                  (cArqTrb)->E1_VENCREA  ,;
                                  (cArqTrb)->E1_VALOR    ,;
                                  (cArqTrb)->E1_HIST})

    (cArqTrb)->(dbSkip())

End

If !Empty(oExcel:aWorkSheet)

    oExcel:Activate()
    oExcel:GetXMLFile(cArquivo)
 
    CpyS2T("\SYSTEM\"+cArquivo, cPath)

    oExcelApp := MsExcel():New()
    oExcelApp:WorkBooks:Open(cPath + "\" + cArquivo) // Abre a planilha
	oExcelApp:SetVisible(.T.)
	
EndIf

If _oPlan <> Nil
	_oPlan:Delete()
	_oPlan := Nil
Endif
    
Return(.T.)