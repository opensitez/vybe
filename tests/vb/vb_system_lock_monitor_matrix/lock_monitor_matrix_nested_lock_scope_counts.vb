' vybe-test: vb/vb_system_lock_monitor_matrix/lock_monitor_matrix_nested_lock_scope_counts
' origin: languages/vb/tests/vb/test_vb_system_lock_monitor_matrix.rs

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
    Sub Main()
        Dim lockObject As New Object()
        Dim value As Integer = 0

        SyncLock lockObject
            value = 1
            SyncLock lockObject
                value = value + 1
            End SyncLock
            value = value * 2
        End SyncLock

        __Check(CStr(value), "4")
    End Sub
End Module
