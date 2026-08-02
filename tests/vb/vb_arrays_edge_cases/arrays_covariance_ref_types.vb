' vybe-test: vb/vb_arrays_edge_cases/arrays_covariance_ref_types
' origin: languages/vb/tests/vb/test_vb_arrays_edge_cases.rs

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

Class Animal
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim dogs(2) As Dog
        Dim animals() As Animal = dogs
        
        __Check(CStr(animals.Length), "3")
        
        Try
            ' This fails at runtime (ArrayTypeMismatchException)
            animals(0) = New Animal()
        Catch ex As System.ArrayTypeMismatchException
            __Check(CStr("Mismatch"), "Mismatch")
        End Try
    End Sub
End Module
