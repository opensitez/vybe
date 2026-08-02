' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_class_static_field_per_type_instantiation
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

Class Counter(Of T)
    Public Shared Count As Integer = 0
End Class

Module Program
    Sub Main()
        Counter(Of String).Count = 100
        Counter(Of Integer).Count = 200
        __Check(CStr(Counter(Of String).Count & "|" & Counter(Of Integer).Count), "100|200")
    End Sub
End Module
