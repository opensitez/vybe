' vybe-test: vb/vb_synclock/synclock_nested
' origin: languages/vb/tests/vb/test_vb_synclock.rs

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

Module M
    Private _lockA As New Object()
    Private _lockB As New Object()
    
    Sub Main()
        SyncLock _lockA
            __Check(CStr("Lock A acquired"), "Lock A acquired")
            SyncLock _lockB
                __Check(CStr("Lock B acquired"), "Lock B acquired")
            End SyncLock
        End SyncLock
    End Sub
End Module
