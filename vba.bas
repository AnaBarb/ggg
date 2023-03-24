Attribute VB_Name = "Módulo1"
Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    'Filtrar Aeroporo na Segmentação de Dados
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Aeroporto1"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Aeroporto].&[SBBE]")
    'Filtrar Pesquisador na Segmentação de Dados
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[Ágata Antônia]")
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[Alike Barbosa]")
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Pesquisador2"). _
        VisibleSlicerItemsList = Array( _
        "[Base_Meta].[Pesquisador].&[Andrea da Silva]")
        
    'Limpar Filtro da Segmentação de Dados
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Aeroporto1").ClearManualFilter 'Aeroporto
    ActiveWorkbook.SlicerCaches("SegmentaçãodeDados_Pesquisador2").ClearManualFilter 'Pesquisador
    
    
End Sub


'
' Macro1 Macro
'

Sub Macro2()

    Dim pesquisador As String
    Dim l As Byte 'número inteiro de 0 a 255
    Dim cont As Byte
        
    l = 2 'linha
    cont = 0 'contador
    
    'Realizar o Loop até a célula ser vazia
    'Cells(linha, coluna)
    Do Until Cells(l, 2) = ""
        pesquisadorana = cont + Cells(l, 2)
        l = l + 1
    Loop
    

End Sub

