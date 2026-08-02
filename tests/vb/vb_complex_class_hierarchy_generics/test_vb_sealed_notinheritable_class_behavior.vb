' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_sealed_notinheritable_class_behavior
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

NotInheritable Class FinalConfig
    Public Property Version As Integer = 1
End Class

Module Program
    Sub Main()
        Dim fc As New FinalConfig()
        __Check(CStr(fc.Version), "1")
    End Sub
End Module
