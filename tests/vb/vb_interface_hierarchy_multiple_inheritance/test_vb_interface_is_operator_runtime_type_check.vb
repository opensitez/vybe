' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_is_operator_runtime_type_check
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface ISupportFastSearch
End Interface

Class FastDatabase
    Implements ISupportFastSearch
End Class

Class SlowDatabase
End Class

Module Program
    Sub Main()
        Dim db1 As Object = New FastDatabase()
        Dim db2 As Object = New SlowDatabase()
        __Check(CStr((TypeOf db1 Is ISupportFastSearch) & "|" & (TypeOf db2 Is ISupportFastSearch)), "True|False")
    End Sub
End Module
