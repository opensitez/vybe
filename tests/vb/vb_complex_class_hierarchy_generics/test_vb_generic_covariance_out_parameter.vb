' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_covariance_out_parameter
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Interface IReadOnlyRepository(Of Out T)
    Function GetFirst() As T
End Interface

Class StringRepository
    Implements IReadOnlyRepository(Of String)
    Public Function GetFirst() As String Implements IReadOnlyRepository(Of String).GetFirst
        Return "CovariantResult"
    End Function
End Class

Module Program
    Sub Main()
        Dim repo As IReadOnlyRepository(Of String) = New StringRepository()
        Dim objRepo As IReadOnlyRepository(Of Object) = repo ' Covariant assignment!
        __Check(CStr(objRepo.GetFirst().ToString()), "CovariantResult")
    End Sub
End Module
