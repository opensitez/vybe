' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_method_calling_base_method
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

Class Parent
    Public Overridable Function Log(msg As String) As String
        Return "Parent: " & msg
    End Function
End Class

Class Child
    Inherits Parent
    Public NotOverridable Overrides Function Log(msg As String) As String
        Return MyBase.Log(msg) & " (Child Sealed)"
    End Function
End Class

Module Program
    Sub Main()
        Dim p As Parent = New Child()
        __Check(CStr(p.Log("Message")), "Parent: Message (Child Sealed)")
    End Sub
End Module
