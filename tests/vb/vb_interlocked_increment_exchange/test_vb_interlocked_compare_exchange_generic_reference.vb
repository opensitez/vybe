' vybe-test: vb/vb_interlocked_increment_exchange/test_vb_interlocked_compare_exchange_generic_reference
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
    Private Function CompareExchangeGeneric(Of T As Class)(ByRef location As T, value As T, comparand As T) As T
        Return Interlocked.CompareExchange(location, value, comparand)
    End Function

    Sub Main()
        Dim s1 As String = "Alpha"
        Dim s2 As String = "Beta"
        Dim target As String = s1
        Dim oldVal = CompareExchangeGeneric(target, s2, s1)
        __Check(CStr(oldVal & "|" & target), "Alpha|Beta")
    End Sub
End Module
