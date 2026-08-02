' vybe-test: vb/vb_object_initializer_edges/object_initializer_with
' origin: languages/vb/tests/vb/test_vb_object_initializer_edges.rs

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
    Public Property Name As String
    Public Property Age As Integer
End Class

Module M
    Sub Main()
        ' With block object initialization
        Dim p1 As New Person With {
            .Name = "Alice",
            .Age = 30
        }
        
        ' Anonymous type with inferred names from properties
        Dim p2 = New With { p1.Name, p1.Age }
        
        __Check(CStr(p1.Name), "Alice")
        __Check(CStr(p2.Age), "30")
    End Sub
End Module
