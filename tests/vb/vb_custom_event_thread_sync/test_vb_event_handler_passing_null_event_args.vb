' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_passing_null_event_args
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

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

Class NullArgsPublisher
    Public Event RawEvent As EventHandler
    Public Sub FireNullArgs()
        RaiseEvent RawEvent(Me, Nothing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim nap As New NullArgsPublisher()
        AddHandler nap.RawEvent, Sub(s, e) __Check(CStr(e Is Nothing), "True")
        nap.FireNullArgs()
    End Sub
End Module
