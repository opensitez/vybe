' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_chain_unsubscribes_all
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

Class Emitter
    Public Event Data As Action(Of Integer)
    Public Sub Push(v As Integer)
        RaiseEvent Data(v)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim h1 As Action(Of Integer) = Sub(v) __Check(CStr("H1:" & v), "H1:1")
        Dim h2 As Action(Of Integer) = Sub(v) __Check(CStr("H2:" & v), "H2:1")

        AddHandler e.Data, h1
        AddHandler e.Data, h2
        e.Push(1)
        RemoveHandler e.Data, h1
        RemoveHandler e.Data, h2
        e.Push(2)
        __Check(CStr("Done"), "Done")
    End Sub
End Module
