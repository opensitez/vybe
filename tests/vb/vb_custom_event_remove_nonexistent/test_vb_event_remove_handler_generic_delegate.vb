' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_generic_delegate
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Class GenericEmitter(Of T)
    Public Event ValueProcessed As EventHandler(Of T)
    Public Sub Process(v As T)
        RaiseEvent ValueProcessed(Me, v)
    End Sub
End Class

Module Program
    Private Sub OnProcess(sender As Object, e As Integer)
        __Check(CStr("Processed: " & e), "Processed: 42")
    End Sub

    Sub Main()
        Dim ge As New GenericEmitter(Of Integer)()
        AddHandler ge.ValueProcessed, AddressOf OnProcess
        ge.Process(42)
        RemoveHandler ge.ValueProcessed, AddressOf OnProcess
        ge.Process(100)
    End Sub
End Module
