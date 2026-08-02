' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notinheritable_type_typeof_checking
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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

NotInheritable Class SealedType
End Class

Module Program
    Sub Main()
        Dim t = GetType(SealedType)
        __Check(CStr(t.IsSealed), "True")
    End Sub
End Module
