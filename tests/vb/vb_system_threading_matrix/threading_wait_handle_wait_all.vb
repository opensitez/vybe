' vybe-test: vb/vb_system_threading_matrix/threading_wait_handle_wait_all
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
        Dim a As New AutoResetEvent(False)
        Dim b As New AutoResetEvent(False)

        a.Set()
        b.Set()

        Dim allDone As WaitHandle() = {a, b}
        __Check(CStr(WaitHandle.WaitAll(allDone)), "True")
    End Sub
End Module
