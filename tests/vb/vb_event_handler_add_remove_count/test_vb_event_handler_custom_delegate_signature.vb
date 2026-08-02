' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_custom_delegate_signature
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

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

Delegate Sub CustomStatusHandler(code As Integer, message As String)

Class StatusNotifier
    Public Event StatusReport As CustomStatusHandler
    Public Sub Notify(c As Integer, m As String)
        RaiseEvent StatusReport(c, m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New StatusNotifier()
        AddHandler n.StatusReport, Sub(c, m) __Check(CStr(c & ": " & m), "200: OK")
        n.Notify(200, "OK")
    End Sub
End Module
