Private Sub Worksheet_Change(ByVal Target As Range)

    Dim rngData As Range
    Dim rngIPN As Range

    Dim interData As Range
    Dim interIPN As Range

    Dim cel As Range

    On Error GoTo Fim

    Application.EnableEvents = False
    Application.ScreenUpdating = False


    Set rngData = Me.Range("H2:H1048576")
Set rngIPN = Me.Range("P2:P1048576")

    Set interData = Intersect(Target, rngData)
    Set interIPN = Intersect(Target, rngIPN)

    If Not interData Is Nothing Then

        For Each cel In interData

            If IsDate(cel.Value) Then

                Me.Cells(cel.Row, "I").Value = _
                    DiaSemanaHoraExtra(cel.Value)

            Else

                Me.Cells(cel.Row, "I").ClearContents

            End If

        Next cel

    End If

    If Not interIPN Is Nothing Then

        For Each cel In interIPN

            If Trim(cel.Value) = "" Then

                Me.Cells(cel.Row, "R").ClearContents
                Me.Cells(cel.Row, "S").ClearContents
                Me.Cells(cel.Row, "T").ClearContents
                Me.Cells(cel.Row, "J").ClearContents

            Else

                BuscarColaborador cel.Row

            End If

        Next cel

    End If
    
    Call AtualizarConflitos

Fim:

    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
