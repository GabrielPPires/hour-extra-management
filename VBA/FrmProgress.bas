Public Sub AtualizarBarra( _
    ByVal Percentual As Double, _
    ByVal Mensagem As String)

    Me.lblTexto.Caption = Mensagem

    Me.lblPercentual.Caption = _
        Format(Percentual, "0") & "%"

    Me.lblBarra.Width = _
        (Percentual / 100) * Me.fraBarra.Width

    DoEvents

End Sub

Private Sub UserForm_Click()

End Sub
