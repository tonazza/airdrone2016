Public Class Popup_tipo_prova

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub


    Private Sub PopUp_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        Select Case Form1.cb_tipo_prova.Text

            Case "Mandata"
                PictureBox1.Image = Camera_Aeraulica.My.Resources.Resources.mandata
                Label1.Text = "Aprire il rubinetto MANDATA e chiudere il rubinetto ASPIRAZIONE"
            Case ("Aspirazione")
                PictureBox1.Image = Camera_Aeraulica.My.Resources.Resources.aspirazione
                Label1.Text = "Aprire il rubinetto ASPIRAZIONE e chiudere il rubinetto MANDATA"
            Case Else
        End Select

    End Sub

End Class