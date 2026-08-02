' vybe-test: vb/vb_system_interlocked_atomic_matrix/interlocked_atomic_matrix_compare_exchange_success_and_fail
' origin: languages/vb/tests/vb/test_vb_system_interlocked_atomic_matrix.rs

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

Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 10

        Dim ok As Integer = Interlocked.CompareExchange(value, 20, 10)
        Dim fail As Integer = Interlocked.CompareExchange(value, 30, 99)

        __Check(CStr(ok), "10")
        __Check(CStr(fail), "20")
        __Check(CStr(value), "20")
    End Sub
End Module
