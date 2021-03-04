Attribute VB_Name = "Módulo1"
Sub inserirautofiltro()
Attribute inserirautofiltro.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Range("A2:D2").Select
    Selection.AutoFilter
    Range("A2").Select
    
End Sub
Sub removereinserirautofiltro()
Attribute removereinserirautofiltro.VB_ProcData.VB_Invoke_Func = " \n14"
    
    Range("A2:D2").Select
    
    If ActiveSheet.AutoFilterMode = True Then
        Selection.AutoFilter
        Range("A2:D2").Select
        Selection.AutoFilter
    Else
        Range("A2:D2").Select
        Selection.AutoFilter
    End If

    Range("A2").Select

End Sub
Sub Retângulo2_Clique()
Range("A2").Select
End Sub
