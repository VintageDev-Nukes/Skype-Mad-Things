Imports System.Net

Public Class Form1

    Public WithEvents oSkype As SKYPE4COMLib.Skype = New SKYPE4COMLib.Skype
    Dim WithEvents tomer As New System.Windows.Forms.Timer With {.Interval = 2500}

    Private Sub Esperar(ByVal interval As Integer)
        Dim sw As New Stopwatch
        sw.Start()
        Do While sw.ElapsedMilliseconds < interval
            ' Allows UI to remain responsive
            Application.DoEvents()
        Loop
        sw.Stop()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles tomer.Tick
        Dim llamada As SKYPE4COMLib.Call = oSkype.PlaceCall(TextBox1.Text)
        Esperar(1000)
        llamada.Finish()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        tomer.Start()
    End Sub
End Class
