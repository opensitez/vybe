' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_lambda_without_reference_no_op
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class Emitter
    Public Event Message As Action(Of String)
    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        AddHandler e.Message, Sub(m) Console.WriteLine("Msg1: " & m)
        ' Attempt to remove a different lambda instance with identical body
        RemoveHandler e.Message, Sub(m) Console.WriteLine("Msg1: " & m)
        e.Dispatch("Test")
    End Sub
End Module
