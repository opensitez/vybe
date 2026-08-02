' vybe-test: vb/vb_class/case_insensitive_field_and_method_access
' origin: languages/vb/tests/vb/vb_class_test.rs

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
    Class Person
        Public Name As String
        Public Function Greet() As String
            Return "Hi " & Name
        End Function
    End Class

    Sub Main()
        Dim P As New Person()
        p.name = "Bob"
        __Check(CStr(p.NAME), "Bob")
        __Check(CStr(p.greet()), "Hi Bob")
    End Sub
End Module
