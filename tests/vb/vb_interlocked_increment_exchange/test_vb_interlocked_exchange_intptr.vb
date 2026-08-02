' vybe-test: vb/vb_interlocked_increment_exchange/test_vb_interlocked_exchange_intptr
' origin: languages/vb/tests/vb/test_vb_interlocked_increment_exchange.rs

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

Module Program
    Sub Main()
        Dim ptr1 As New IntPtr(1000)
        Dim ptr2 As New IntPtr(2000)
        Dim oldPtr = Interlocked.Exchange(ptr1, ptr2)
        __Check(CStr(oldPtr.ToInt32() & "|" & ptr1.ToInt32()), "1000|2000")
    End Sub
End Module
