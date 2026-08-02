' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_nested_class_properties
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

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

Module Program
    Class Engine
        Public Property Core As CoreUnit
    End Class

    Class CoreUnit
        Public Property Code As String = "C100"
    End Class

    Sub Main()
        Dim eng As New Engine()
        __Check(CStr(eng IsNot Nothing AndAlso eng.Core IsNot Nothing), "False")
    End Sub
End Module
