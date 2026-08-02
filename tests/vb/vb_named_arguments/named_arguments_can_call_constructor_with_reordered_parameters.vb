' vybe-test: vb/vb_named_arguments/named_arguments_can_call_constructor_with_reordered_parameters
' origin: languages/vb/tests/vb/test_vb_named_arguments.rs

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

Class Person
    Public Name As String
    Public Age As Integer

    Public Sub New(age As Integer, name As String)
        Me.Name = name
        Me.Age = age
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Person(name:="Gus", age:=41)
        __Check(CStr(p.Name), "Gus")
        __Check(CStr(p.Age), "41")
    End Sub
End Module
