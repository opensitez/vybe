' vybe-test: vb/vb_array_true_for_all_exists/test_vb_array_trueforall_complex_object_properties
' origin: languages/vb/tests/vb/test_vb_array_true_for_all_exists.rs

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

Class Account
    Public Balance As Decimal
    Public Sub New(b As Decimal)
        Balance = b
    End Sub
End Class

Module Program
    Sub Main()
        Dim accs As Account() = {New Account(100D), New Account(250D), New Account(50D)}
        Dim allPositive As Boolean = Array.TrueForAll(accs, Function(a) a.Balance > 0D)
        Dim allWealthy As Boolean = Array.TrueForAll(accs, Function(a) a.Balance >= 200D)
        __Check(CStr(allPositive & "|" & allWealthy), "True|False")
    End Sub
End Module
