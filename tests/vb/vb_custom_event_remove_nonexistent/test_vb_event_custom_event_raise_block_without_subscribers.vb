' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_custom_event_raise_block_without_subscribers
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

Class SafeCustomEvent
    Public Custom Event EventTest As Action
        AddHandler(value As Action) : End AddHandler
        RemoveHandler(value As Action) : End RemoveHandler
        RaiseEvent()
            __Check(CStr("RaiseBlock executed directly"), "RaiseBlock executed directly")
        End RaiseEvent
    End Event
    Public Sub Fire()
        RaiseEvent EventTest()
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New SafeCustomEvent()
        s.Fire()
    End Sub
End Module
