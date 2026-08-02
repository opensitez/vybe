' vybe-test: vb/vb_system_threading_atomic/system_threading_atomic_exchange_and_compare_exchange
' origin: languages/vb/tests/vb/test_vb_system_threading_atomic.rs

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
        Dim exchangeResult As Integer = Interlocked.Exchange(value, 20)
        __Check(CStr(exchangeResult), "10")
        __Check(CStr(value), "20")

        Dim failedCompare As Integer = Interlocked.CompareExchange(value, 30, 5)
        __Check(CStr(failedCompare), "20")
        __Check(CStr(value), "20")

        Dim successCompare As Integer = Interlocked.CompareExchange(value, 30, 20)
        __Check(CStr(successCompare), "20")
        __Check(CStr(value), "30")
    End Sub
End Module
