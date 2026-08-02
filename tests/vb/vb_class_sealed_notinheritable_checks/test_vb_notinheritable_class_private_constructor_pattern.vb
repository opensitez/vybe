' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notinheritable_class_private_constructor_pattern
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

NotInheritable Class Singleton
    Public Shared ReadOnly Instance As New Singleton()
    Private Sub New()
    End Sub
    Public Function GetName() As String
        Return "SingletonInstance"
    End Function
End Class

Module Program
    Sub Main()
        __Check(CStr(Singleton.Instance.GetName()), "SingletonInstance")
    End Sub
End Module
