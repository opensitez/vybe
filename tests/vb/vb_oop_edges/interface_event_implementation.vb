' vybe-test: vb/vb_oop_edges/interface_event_implementation
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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

Interface INotify
    Event Raised()
End Interface

Class Notifier
    Implements INotify
    
    Public Event Raised() Implements INotify.Raised
    
    Public Sub Trigger()
        RaiseEvent Raised()
    End Sub
End Class

Module M
    Sub Main()
        Dim n As New Notifier()
        AddHandler n.Raised, Sub() __Check(CStr("Event Triggered"), "Event Triggered")
        n.Trigger()
    End Sub
End Module
