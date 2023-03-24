Attribute VB_Name = "M�dulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    'Filtrar Aeroporo na Segmenta��o de Dados
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Aeroporto1"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Aeroporto].&[SBBE]")
    'Filtrar Pesquisador na Segmenta��o de Dados
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[�gata Ant�nia]")
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[Alike Barbosa]")
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[Andrea da Silva]")
        
    'Limpar Filtro da Segmenta��o de Dados
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Aeroporto1").ClearManualFilter 'Aeroporto
    ActiveWorkbook.SlicerCaches("Segmenta��odeDados_Pesquisador2").ClearManualFilter 'Pesquisador
    
    
End Sub


'
' Macro1 Macro
'

Sub Macro2()

    Dim pesquisador As String
    Dim l As Byte 'n�mero inteiro de 0 a 255
    Dim cont As Byte
        
    l = 2 'linha
    cont = 0 'contador
    
    'Realizar o Loop at� a c�lula ser vazia
    'Cells(linha, coluna)
    Do Until Cells(l, 2) = ""
        pesquisadorana = cont + Cells(l, 2)
        l = l + 1
    Loop
    

End Sub

