' vybe-test: vb/vb_system_threading_matrix/threading_interlocked_exchange_and_compare_exchange
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
        Dim value As Integer = 5
        Dim old As Integer = Interlocked.Exchange(value, 9)

        Dim matched As Integer = Interlocked.CompareExchange(value, 11, 9)
        Dim unmatched As Integer = Interlocked.CompareExchange(value, 13, 5)

        __Check(CStr(old), "5")
        __Check(CStr(matched), "9")
        __Check(CStr(unmatched), "11")
        __Check(CStr(value), "11")
    End Sub
End Module
