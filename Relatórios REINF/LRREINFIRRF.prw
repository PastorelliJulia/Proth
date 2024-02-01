// //#INCLUDE "TOPCONN.CH"
// //#INCLUDE "PROTHEUS.CH"
// #INCLUDE "RWMAKE.CH"
// //#INCLUDE "MSOBJECTS.CH" 

// ' //--------------------------------------------------------------------------------------------------------------//
// ' //------------------|      Data       |---|      Função      |---|        Autor        |------------------------//
// ' //------------------|     26/12/23    |---|    LRREINFINSS   |---|   Julia Pastorelli  |------------------------//
// ' //--------------------------------------------------------------------------------------------------------------//
// ' // Descrição: Relatorio conferencia REINF - INSS                                                                //			
// ' //--------------------------------------------------------------------------------------------------------------//

// ' User Function ()

// ' 	Local cDesc1        := "Este programa tem como objetivo imprimir relatorio "
// ' 	Local cDesc2        := "de REINF da taxa INSS"
// ' 	Local titulo       	:= "REINF INSS"
// ' 	Local cPerg	  		:= "LRREINF"
// ' 	Local aOrd 			:= {}
// ' 	Local aArea        	:= GetArea()
// ' 	Private tamanho     := "G"
// ' 	Private nomeprog    := "REINF INSS" 
// ' 	Private nTipo       := 15
// ' 	Private aReturn     := { "Zebrado", 1, "Administracao", 1, 2, 1, "", 1}
// ' 	Private nLastKey    := 0
// ' 	Private wnrel      	:= "REINF_INSS" 
// ' 	Private aRotina := {}
// ' 	Private oApoio	:= LibApoio():New()

// ' 	pergunte(cPerg,.f.)

// ' 	wnrel := SetPrint("",NomeProg,cPerg,@titulo,cDesc1,cDesc2,,.F.,aOrd,.F.,Tamanho,,.F.)

// ' 	If nLastKey == 27
// ' 		Return
// ' 	Endif

// ' 	SetDefault(aReturn,"")
// ' 	If nLastKey == 27
// ' 		Return
// ' 	Endif

// ' 	nTipo := If(aReturn[4]==1,15,18)

// ' 	fGeraExcel()
// ' 	RestArea(aArea)
// ' RETURN

// ' //--------------------------------------------------------------------------------------------------------------//
// ' //| Gera excell
// ' //--------------------------------------------------------------------------------------------------------------//
// ' Static Function fGeraExcel()
 
// ' 	Local cQuery 	:= ""
// ' 	Local cPath     := ""
// '     Local cNameFile := ""
// ' 	Local aRec	:= ""
// ' 	Local aCab	:= {}
// ' 	LOCAL x		:= 0
// '     Private lMsErroAuto := .F.
// '     Private oExcel  := FWMSEXCEL():New()

// ' 	cQuery := " SELECT  "+CRLF
// ' 	cQuery += " F1_FILIAL,F1_ESPECIE,F1_DOC,F1_SERIE,F1_FORNECE,A2_NOME,F1_EMISSAO,B1_DESC,D1_CODISS,D1_ALIQINS,D1_VALINS,F1_INSS,E2_NATUREZ,E2_VALLIQ,E2_SALDO"
// ' 	// cQuery += " F1_FILIAL,F1_ESPECIE,F1_DOC,F1_SERIE,F1_FORNECE,A2_NOME,-A2_TIPO-,-A2_MUN-,F1_EMISSAO,-D1_ITEM-,D1_COD,B1_DESC,D1_TES,-D1_QUANT-,D1_CODISS,X5_DESCRI,D1_BASEINS,D1_ALIQINS,D1_VALINS,F1_INSS,E2_NATUREZ,E2_VALLIQ,E2_SALDO "+CRLF

// ' 	cQuery += " FROM " + RetSqlName("SF1") +  " (NOLOCK) " +CRLF
// ' 	cQuery += " INNER JOIN " + RetSqlName("SA2") +  "  (NOLOCK) "+CRLF 
// ' 	cQuery += " ON A2_FILIAL = F1_FILIAL AND A2_COD = F1_FORNECE AND A2_LOJA = F1_LOJA AND " + RetSqlName("SA2") +  ".D_E_L_E_T_ = '' "+CRLF 
// ' 	cQuery += " INNER JOIN " + RetSqlName("SD1") +  "  (NOLOCK) "+CRLF
// ' 	cQuery += " ON D1_FILIAL = F1_FILIAL AND D1_DOC = F1_DOC AND D1_SERIE = F1_SERIE AND D1_FORNECE = F1_FORNECE AND D1_LOJA = F1_LOJA AND " + RetSqlName("SD1") +  ".D_E_L_E_T_ = ''" + CRLF

// ' 	cQuery += " INNER JOIN " + RetSqlName("SB1") +  "  (NOLOCK) "+CRLF
// ' 	cQuery += " ON B1_FILIAL = D1_FILIAL AND B1_COD = D1_COD AND " + RetSqlName("SB1") +  ".D_E_L_E_T_ = ''" + CRLF

// ' 	cQuery += " INNER JOIN " + RetSqlName("SE2") +  "  (NOLOCK) "+CRLF
// ' 	cQuery += " ON E2_FILIAL = F1_FILIAL AND E2_NUM = F1_DOC AND E2_FORNECE = F1_FORNECE AND " + RetSqlName("SE2") +  ".D_E_L_E_T_ = ''" + CRLF
// ' 	cQuery += " INNER JOIN " + RetSqlName("SED") +  "  (NOLOCK) "+CRLF
// ' 	cQuery += " ON ED_CODIGO = E2_NATUREZ AND " + RetSqlName("SED") +  ".D_E_L_E_T_ = ''" + CRLF

// ' 	cQuery += " WHERE "+ CRLF
// '     cQuery += " F1_ESPECIE IN ('CF','CTE','NF','NFCEE','NFFA','NFS','NTST','RCLOC','SPED') " + CRLF
// ' 	cQuery += " F1_FILIAL = 'A1' "
// ' 	cQuery += " AND F1_EMISSAO BETWEEN '" +DTOS(mv_par01)+ "' AND '" +DTOS(mv_par02)+ "'" + CRLF
// ' 	cQuery += " AND " + RetSqlName("SF1") +  ".D_E_L_E_T_ =''"

// ' 	aRec := U_QryArr(cQuery) 
// '         aadd( aCab , {"FILIAL"  				,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"ESPECIE"  				,  1 , 1 , 0 } )
// ' 	    aadd( aCab , {"DOCUMENTO"				,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"SERIE"					,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"COD FORNECEDOR" 			    ,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"FORNECEDOR" 			    ,  1 , 1 , 0 } )
// '         aadd( aCab , {"EMISSAO"	               	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"DESCRICAO"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"COD ISS"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"BASE INSS ITEM"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"ALIQ INSS ITEM"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"TOT INSS NF"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"NATUREZA"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"VAL LIQ"             	,  1 , 1 , 0 } )
// ' 		aadd( aCab , {"SALDO"             	,  1 , 1 , 0 } )

// ' 		oExcel:AddworkSheet("REINF INSS")
// ' 	    oExcel:AddTable ("REINF INSS","REINF INSS")
// ' 		for x:=1 to len(aCab)
// ' 		                   //(< cWorkSheet >, < cTable >, < cColumn >, < nAlign >, < nFormat >, < lTotal >)
// ' 		    oExcel:AddColumn("REINF INSS","REINF INSS",aCab[x][1],aCab[x][3],aCab[x][2])
// ' 	    next x

// ' 		for x := 1 to len(aRec)
// ' 		    oExcel:AddRow("REINF INSS","REINF INSS",{aRec[x][1],aRec[x][2],aRec[x][3],aRec[x][4],aRec[x][5],aRec[x][6],aRec[x][7],aRec[x][8],sTod(aRec[x][9]),aRec[x][10],aRec[x][11],aRec[x][12],aRec[x][13],aRec[x][14],aRec[x][15]})
// ' 	    next x
// '         oExcel:Activate()
// ' 		cPath := AllTrim(GetTempPath())
// '         cNameFile := cPath + cValToChar(Randomize( 1, 1000 )) + "_REINFINSS.xls"

// ' 	    oExcel:GetXMLFile(cNameFile)
// ' 		ShellExecute( "Open",cNameFile, '', '', 1 ) 

// ' Return()  
