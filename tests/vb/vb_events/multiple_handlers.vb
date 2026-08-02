' vybe-test: vb/vb_events/multiple_handlers
' origin: languages/vb/tests/vb/test_vb_events.rs

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

Class Notifier
    Public Event Notify()
    Public Sub Fire()
        RaiseEvent Notify()
    End Sub
End Class

Module M
    Sub Handler1()
        __Check(CStr("h1"), "h1")
    End Sub
    Sub Handler2()
        __Check(CStr("h2"), "h2")
    End Sub
    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Notify, AddressOf Handler1
        AddHandler n.Notify, AddressOf Handler2
        n.Fire()
    End Sub
End Module
