Imports System.Security.Cryptography

Module Module1
    Public Nombre(6) As String
    Public CUI(6) As String
    Public Direccion(6) As String
    Public Plazo(6) As String
    Public Aparato(6) As String
    Public Parcial(6) As Integer
    Public Descuento(6) As Double
    Public Total(6) As Double
    Public Fila As Byte


    Public Const TV_C = 250
    Public Const TV_L = 300
    Public Const TF_C = 550
    Public Const TF_L = 600
    Public Const LT_C = 770
    Public Const LT_L = 800
    Public Const RG_C = 1000
    Public Const RG_L = 1200

    Sub ACEPTAR()
        If Fila <= 6 Then
            If (Form1.TextBox1.Text <> "") Then
                Nombre(Fila) = Form1.TextBox1.Text
            Else
                MsgBox("Error, no ingresó nombre")
                Form1.TextBox1.Focus() : Exit Sub
            End If

            If (Form1.TextBox2.Text <> "") Then
                Direccion(Fila) = Form1.TextBox2.Text
            Else
                MsgBox("Error, no ingresó direccion")
                Form1.TextBox2.Focus() : Exit Sub
            End If

            If (Form1.RadioButton1.Checked) Then
                Plazo(Fila) = "Corto"
            ElseIf (Form1.RadioButton2.Checked) Then
                Plazo(Fila) = "Largo"
            Else MsgBox("No seleccionó el plazo que desea")
            End If


            Select Case (Form1.ComboBox1.SelectedIndex)
                Case 0 : Aparato(Fila) = "TV"
                Case 1 : Aparato(Fila) = "TELÉFONO"
                Case 2 : Aparato(Fila) = "LAPTOP"
                Case 3 : Aparato(Fila) = "REFRIGERADORA"
                Case Else
                    MsgBox("No seleccionó tipo de aparato") : Exit Sub
            End Select

            Select Case (Form1.ComboBox1.SelectedIndex)
                Case 0 And (Form1.RadioButton1.Checked) : Parcial(Fila) = TV_C
                Case 1 And (Form1.RadioButton1.Checked) : Parcial(Fila) = TF_C
                Case 2 And (Form1.RadioButton1.Checked) : Parcial(Fila) = LT_C
                Case 3 And (Form1.RadioButton1.Checked) : Parcial(Fila) = RG_C

                Case 0 And (Form1.RadioButton2.Checked) : Parcial(Fila) = TV_L
                Case 1 And (Form1.RadioButton2.Checked) : Parcial(Fila) = TF_L
                Case 2 And (Form1.RadioButton2.Checked) : Parcial(Fila) = LT_L
                Case 3 And (Form1.RadioButton2.Checked) : Parcial(Fila) = RG_L
                Case Else
                    MsgBox("No seleccionó tipo de aparato o plazo") : Exit Sub
            End Select

            Select Case (Form1.ComboBox1.SelectedIndex)
                Case 0 And (Form1.RadioButton1.Checked) : Descuento(Fila) = Parcial(Fila) * 0.15
                Case 1 And (Form1.RadioButton1.Checked) : Descuento(Fila) = Parcial(Fila) * 0.1
                Case 2 And (Form1.RadioButton1.Checked) : Descuento(Fila) = Parcial(Fila) * 0.1
                Case 3 And (Form1.RadioButton1.Checked) : Descuento(Fila) = Parcial(Fila) * 0.15

                Case 0 And (Form1.RadioButton2.Checked) : Descuento(Fila) = Parcial(Fila) * 0.05
                Case 1 And (Form1.RadioButton2.Checked) : Descuento(Fila) = Parcial(Fila) * 0.05
                Case 2 And (Form1.RadioButton2.Checked) : Descuento(Fila) = Parcial(Fila) * 0.05
                Case 3 And (Form1.RadioButton2.Checked) : Descuento(Fila) = Parcial(Fila) * 0.05
            End Select
        End If
        Total(Fila) = Parcial(Fila) - Descuento(Fila)
        If (Fila = 7) Then
            MsgBox("Vectores llenos")
        End If
    End Sub
    Sub limpiar_entradas()
        Form1.TextBox1.Clear()
        Form1.RadioButton1.Checked = False
        Form1.RadioButton2.Checked = False
        Form1.ComboBox1.SelectedIndex = -1
        Form1.TextBox1.Focus()
    End Sub

    Sub limpiar_grid()
        Array.Clear(Nombre, 0, Nombre.Length)
        Array.Clear(Direccion, 0, Direccion.Length)
        Array.Clear(CUI, 0, CUI.Length)
        Array.Clear(Plazo, 0, Plazo.Length)
        Array.Clear(Aparato, 0, Aparato.Length)
        Array.Clear(Parcial, 0, Parcial.Length)
        Array.Clear(Descuento, 0, Descuento.Length)
        Array.Clear(Total, 0, Total.Length)

        Fila = 0
    End Sub
End Module
