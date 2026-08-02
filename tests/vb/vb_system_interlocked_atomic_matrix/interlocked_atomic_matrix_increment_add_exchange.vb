' vybe-test: vb/vb_system_interlocked_atomic_matrix/interlocked_atomic_matrix_increment_add_exchange
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
        Dim value As Integer = 0

        Dim inc1 As Integer = Interlocked.Increment(value)
        Dim add5 As Integer = Interlocked.Add(value, 5)
        Dim prev As Integer = Interlocked.Exchange(value, 42)

        __Check(CStr(inc1), "1")
        __Check(CStr(add5), "6")
        __Check(CStr(prev), "6")
        __Check(CStr(value), "42")
    End Sub
End Module
