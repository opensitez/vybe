' vybe-test: vb/vb_property_expression_bodied/test_vb_property_single_line_getter
' origin: languages/vb/tests/vb/test_vb_property_expression_bodied.rs

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
    Public Property FirstName As String = "John"
    Public Property LastName As String = "Doe"

    Public ReadOnly Property FullName As String
        Get
            Return FirstName & " " & LastName
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        __Check(CStr(p.FullName), "John Doe")
    End Sub
End Module
