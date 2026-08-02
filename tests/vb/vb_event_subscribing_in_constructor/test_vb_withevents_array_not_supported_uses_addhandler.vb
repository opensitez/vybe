' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_array_not_supported_uses_addhandler
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

Imports System

Class Button
    Public Property ID As Integer
    Public Event Click As EventHandler
    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Class FormContainer
    Private buttons As Button()

    Public Sub New()
        buttons = New Button() {New Button With {.ID = 1}, New Button With {.ID = 2}}
        For Each btn In buttons
            AddHandler btn.Click, AddressOf OnButtonClick
        Next
    End Sub

    Private Sub OnButtonClick(sender As Object, e As EventArgs)
        Dim btn = CType(sender, Button)
        Console.WriteLine("Button " & btn.ID & " Clicked")
    End Sub

    Public Sub TestClicks()
        buttons(0).PerformClick()
        buttons(1).PerformClick()
    End Sub
End Class

Module Program
    Sub Main()
        Dim form As New FormContainer()
        form.TestClicks()
    End Sub
End Module
