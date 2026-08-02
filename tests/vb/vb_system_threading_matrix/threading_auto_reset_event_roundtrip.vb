' vybe-test: vb/vb_system_threading_matrix/threading_auto_reset_event_roundtrip
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
        Dim signal As New AutoResetEvent(False)

        __Check(CStr(signal.WaitOne(1)), "False")
        signal.Set()
        __Check(CStr(signal.WaitOne(2000)), "True")
        __Check(CStr(signal.WaitOne(1)), "False")
    End Sub
End Module
