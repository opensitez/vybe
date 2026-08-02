' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_exception_handling_during_raise
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System

Class RobustPublisher
    Public Event ActionEvent As EventHandler

    Public Sub SafeRaise()
        If ActionEventEvent IsNot Nothing Then
            For Each del In ActionEventEvent.GetInvocationList()
                Try
                    del.DynamicInvoke(Me, EventArgs.Empty)
                Catch ex As Exception
                    Console.WriteLine("Handler Error Handled")
                End Try
            Next
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New RobustPublisher()
        AddHandler p.ActionEvent, Sub(s, e) Throw New InvalidOperationException("Handler Fail")
        AddHandler p.ActionEvent, Sub(s, e) Console.WriteLine("Handler 2 Executed")
        p.SafeRaise()
    End Sub
End Module
