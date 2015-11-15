Public Class Popup_dati_incompleti

    Private Sub Popup_dati_incompleti_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Form1.Timer_dati_incompleti.Stop()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

End Class