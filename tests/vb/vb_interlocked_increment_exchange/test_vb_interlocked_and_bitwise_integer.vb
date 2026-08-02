' vybe-test: vb/vb_interlocked_increment_exchange/test_vb_interlocked_and_bitwise_integer
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

Imports System.Threading

Module Program
    Sub Main()
        Dim flags As Integer = 7 ' 0111
        Dim oldVal = Interlocked.And(flags, 3) ' 0011 -> 0011 (3)
        __Check(CStr("Old: " & oldVal & " | New: " & flags), "Old: 7 | New: 3")
    End Sub
End Module
