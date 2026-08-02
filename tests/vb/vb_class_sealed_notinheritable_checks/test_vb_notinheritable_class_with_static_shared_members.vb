' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notinheritable_class_with_static_shared_members
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

NotInheritable Class UtilityHelper
    Public Shared Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function
End Class

Module Program
    Sub Main()
        __Check(CStr(UtilityHelper.Multiply(6, 7)), "42")
    End Sub
End Module
