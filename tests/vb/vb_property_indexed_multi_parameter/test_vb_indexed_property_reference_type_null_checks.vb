' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_reference_type_null_checks
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Class NullableStore
    Private items(1) As String
    Default Public Property Item(i As Integer) As String
        Get
            Return items(i)
        End Get
        Set(val As String)
            items(i) = val
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ns As New NullableStore()
        __Check(CStr(ns(0) Is Nothing), "True")
        ns(0) = "Set"
        __Check(CStr(ns(0) Is Nothing), "False")
    End Sub
End Module
