' vybe-test: vb/vb_system_threading_matrix/threading_monitor_locking_changes_visibility
' origin: languages/vb/tests/vb/test_vb_system_threading_matrix.rs

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

Imports System.Threading

Module M
    Sub Main()
        Dim lockObj As New Object()
        Dim sum As Integer = 0

        SyncLock lockObj
            sum = 1
        End SyncLock

        __Check(CStr(sum = 1), "True")
        __Check(CStr(Monitor.IsEntered(lockObj)), "False")
    End Sub
End Module
