Public Class Form1
    Private Sub ACEPTARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ACEPTARToolStripMenuItem.Click
        ACEPTAR()
    End Sub

    Private Sub MOSTRARVECTORESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MOSTRARVECTORESToolStripMenuItem.Click
        Dim J As Byte

        For J = 0 To 6
            DataGridView1.Rows.Add(Nombre(J), CUI(J), Direccion(J), Plazo(J), Aparato(J), Parcial(J), Descuento(J), Total(J))
        Next J
    End Sub

    Private Sub SALIRToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SALIRToolStripMenuItem.Click
        If (MsgBox("¿Desea Salir?", vbQuestion + vbYesNo, "Mensaje de Salida") = vbYes) Then
            End
        Else
            'Llamada al procedimiento que limpia entradas y limpia listas
            limpiar_entradas()
            limpiar_grid()

        End If
    End Sub

    Private Sub LIMPIARVECTORESToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIMPIARVECTORESToolStripMenuItem.Click
        limpiar_grid()
    End Sub

    Private Sub LIMPIARENTRADASToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LIMPIARENTRADASToolStripMenuItem.Click
        limpiar_entradas()
    End Sub
End Class
