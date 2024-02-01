#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"
#INCLUDE "RWMAKE.CH"
//#INCLUDE "MSOBJECTS.CH"

//--------------------------------------------------------------------------------------------------------------//
//------------------|      Data       |---|      Função      |---|        Autor        |------------------------//
//------------------|     17/04/24    |---|    LRREINFPCC    |---|   Julia Pastorelli  |------------------------//
//--------------------------------------------------------------------------------------------------------------//
// Descrição: Relatorio conferencia REINF - PCC                                                               //			
//--------------------------------------------------------------------------------------------------------------//

User Function LRREINFPCC()

	Local cDesc1        := "Este programa tem como objetivo imprimir relatorio "
	Local cDesc2        := "de REINF da taxa PCC"
	Local titulo       	:= "REINF PCC"
	Local cPerg	  		:= "LRREINF"
	Local aOrd 			:= {}
	Local aArea        	:= GetArea()
	Private tamanho     := "G" 
	Private nomeprog    := "REINF PCC" // Coloque aqui o nome do programa para impressao no cabecalho
	Private nTipo       := 15
	Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
	Private nLastKey    := 0
	Private wnrel      	:= "REINF PCC" // Coloque aqui o nome do arquivo usado para impressao em disco
	Private aRotina := {}
	Private oApoio	:= LibApoio():New()

	pergunte(cPerg,.f.)

	wnrel := SetPrint("",NomeProg,cPerg,@titulo,cDesc1,cDesc2,,.F.,aOrd,.F.,Tamanho,,.F.)

	If nLastKey == 27
		Return
	Endif

	SetDefault(aReturn,"")
	If nLastKey == 27
		Return
	Endif

	nTipo := If(aReturn[4]==1,15,18)

	fGeraExcel()
	RestArea(aArea)
RETURN

//+--------------------------------------------------------------------------------------+
//| Gera excell
//+--------------------------------------------------------------------------------------+
Static Function fGeraExcel()

	Local cQuery 	:= ""
	Local cPath     := ""
    Local cNameFile := ""
	Local aRec	:= ""
	Local aCab	:= {}
	LOCAL x		:= 0
    Private lMsErroAuto := .F.
    Private oExcel  := FWMSEXCEL():New()
	
	cQuery := " SELECT  "+CRLF
	cQuery += " F1_FILIAL,F1_ESPECIE,F1_DOC,E2_PARCELA,F1_SERIE,F1_FORNECE,A2_NOME,E2_EMISSAO,F1_VALBRUT,E2_VALLIQ,E2_SALDO,E2_PIS,E2_COFINS,E2_CSLL "+CRLF
	cQuery += " FROM " + RetSqlName("SF1") +  " (NOLOCK) " +CRLF

	cQuery += " INNER JOIN " + RetSqlName("SA2") +  "  (NOLOCK) "+CRLF
	cQuery += " ON A2_FILIAL = F1_FILIAL AND A2_COD = F1_FORNECE AND A2_LOJA = F1_LOJA AND " + RetSqlName("SA2") + ".D_E_L_E_T_ = '' "+CRLF

    cQuery += " INNER JOIN " + RetSqlName("SE2") +  "  (NOLOCK) "+CRLF 
	cQuery += " ON  E2_FILIAL = F1_FILIAL AND E2_NUM = F1_DOC AND E2_FORNECE = F1_FORNECE AND E2_EMISSAO = F1_EMISSAO AND " + RetSqlName("SE2") + ".D_E_L_E_T_ = '' "+CRLF

	// cQuery += " INNER JOIN " + RetSqlName("SE2") +  "  (NOLOCK) "+CRLF
	// cQuery += " ON E2_FILIAL = F1_FILIAL AND E2_NUM = F1_DOC AND E2_FORNECE = F1_FORNECE AND E2_EMISSAO = F1_EMISSAO AND" + RetSqlName("SE2") +  ".D_E_L_E_T_ = ''" + CRLF

    cQuery += " WHERE "+ CRLF	
    cQuery += " F1_ESPECIE IN ('NFS','CTEOS','CTE') " + CRLF
	cQuery += " AND F1_EMISSAO BETWEEN '" +DTOS(mv_par01)+ "' AND '" +DTOS(mv_par02)+ "'" + CRLF
	cQuery += " AND " + RetSqlName("SF1") +  ".D_E_L_E_T_ =''"

	aRec := U_QryArr(cQuery)

        aadd( aCab , {"FILIAL"  				,  1 , 1 , 0 } )
	    aadd( aCab , {"ESPECIE"			        ,  1 , 1 , 0 } )
        aadd( aCab , {"DOCUMENTO" 			    ,  1 , 1 , 0 } )
        aadd( aCab , {"PARCELA" 				,  1 , 1 , 0 } )
        aadd( aCab , {"SERIE"	               	,  1 , 1 , 0 } )
		aadd( aCab , {"COD FORNECEDOR"          ,  1 , 1 , 0 } )
        aadd( aCab , {"FORNECEDOR"              ,  1 , 1 , 0 } )
        aadd( aCab , {"EMISSAO"                 ,  1 , 1 , 0 } )
		aadd( aCab , {"TOT NF"        	        ,  1 , 1 , 0 } )
        aadd( aCab , {"VAL LIQ"          		,  1 , 1 , 0 } )
		aadd( aCab , {"SALDO"       		    ,  1 , 1 , 0 } )
		aadd( aCab , {" VL PIS"		      		,  1 , 1 , 0 } )
		aadd( aCab , {"VL COFINS"     		  	,  1 , 1 , 0 } )
		aadd( aCab , {"VL CSLL"      			,  1 , 1 , 0 } )


		oExcel:AddworkSheet("REINF PCC")
	    oExcel:AddTable ("REINF PCC","REINF PCC")
		for x:=1 to len(aCab)
		                   //(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)
		    oExcel:AddColumn("REINF PCC","REINF PCC",aCab[x][1],aCab[x][3],aCab[x][2])
	    next x

		for x := 1 to len(aRec)
		    oExcel:AddRow("REINF PCC","REINF PCC",{aRec[x][1],aRec[x][2],aRec[x][3],aRec[x][4],aRec[x][5],aRec[x][6],aRec[x][7],sTod(aRec[x][8]);
												,aRec[x][9],aRec[x][10],aRec[x][11],aRec[x][12],aRec[x][13],aRec[x][14]})
	    next x
        oExcel:Activate()
		cPath := AllTrim(GetTempPath())
        cNameFile := cPath + cValToChar(Randomize( 1, 1000 )) + "_REINFPCC.xls"

	    oExcel:GetXMLFile(cNameFile)
		ShellExecute( "Open",cNameFile, '', '', 1 )

Return()
