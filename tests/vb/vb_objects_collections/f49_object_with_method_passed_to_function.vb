' vybe-test: vb/vb_objects_collections/f49_object_with_method_passed_to_function
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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
    Public Name As String
    Public Function Greet() As String
        Return "Hello, " & Name
    End Function
End Class
Function GetGreeting(g As Greeter) As String
    Return g.Greet()
End Function
Dim gr As New Greeter()
gr.Name = "World"
__Check(CStr(GetGreeting(gr)), "Hello, World")
