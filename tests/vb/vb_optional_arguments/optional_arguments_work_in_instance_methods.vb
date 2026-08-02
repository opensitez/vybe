' vybe-test: vb/vb_optional_arguments/optional_arguments_work_in_instance_methods
' origin: languages/vb/tests/vb/test_vb_optional_arguments.rs

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

Class Greeter
    Public Function Build(name As String, Optional prefix As String = "Hi") As String
        Return prefix & " " & name
    End Function
End Class

Module M
    Sub Main()
        Dim greeter As New Greeter()
        __Check(CStr(greeter.Build("Dana")), "Hi Dana")
        __Check(CStr(greeter.Build("Eli", "Hello")), "Hello Eli")
    End Sub
End Module
