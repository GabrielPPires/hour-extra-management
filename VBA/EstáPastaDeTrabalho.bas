'=================================================
' INICIALIZAÇÃO DO SISTEMA
'
' Garante que os eventos do Excel sejam
' restaurados ao abrir a planilha.
'
' Evita problemas causados por
' Application.EnableEvents = False
' após encerramentos inesperados.
'=================================================

Private Sub Workbook_Open()

    Application.EnableEvents = True

End Sub
