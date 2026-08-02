' vybe-test: vb/vb_property_writeonly_set_semantics/test_vb_property_writeonly_write_side_effects
' origin: languages/vb/tests/vb/test_vb_property_writeonly_set_semantics.rs

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

Class AuditTracker
    Public AuditLog As String = ""

    Public WriteOnly Property LogEntry As String
        Set(value As String)
            AuditLog &= "[" & value & "];"
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim tracker As New AuditTracker()
        tracker.LogEntry = "Event1"
        tracker.LogEntry = "Event2"
        __Check(CStr(tracker.AuditLog), "[Event1];[Event2];")
    End Sub
End Module
