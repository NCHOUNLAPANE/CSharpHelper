'--FORM MOVE--'

Dim x, y As Integer
Dim newpoint As New Point

#Region "Form Move"
    Private Sub frmMain_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown, panelTitle.MouseDown, labelTitle.MouseDown
        x = Control.MousePosition.X - Me.Location.X
        y = Control.MousePosition.Y - Me.Location.Y
    End Sub



    Private Sub frmMain_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove, panelTitle.MouseMove, labelTitle.MouseMove
        If e.Button = MouseButtons.Left Then
            newpoint = Control.MousePosition
            newpoint.X -= x
            newpoint.Y -= y
            Me.Location = newpoint
            Application.DoEvents()
        End If
    End Sub


#End Region

'--END FORM MOVE--'