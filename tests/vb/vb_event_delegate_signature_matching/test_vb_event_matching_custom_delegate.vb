' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_event_matching_custom_delegate
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Delegate Sub StatusChangeHandler(oldStatus As String, newStatus As String)

Class Machine
    Public Event StatusChanged As StatusChangeHandler
    Public Sub UpdateStatus(n As String)
        RaiseEvent StatusChanged("Offline", n)
    End Sub
End Class

Module Program
    Sub Main()
        Dim m As New Machine()
        AddHandler m.StatusChanged, Sub(o, n) __Check(CStr(o & " -> " & n), "Offline -> Online")
        m.UpdateStatus("Online")
    End Sub
End Module
