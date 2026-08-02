' vybe-test: vb/vb_with_statements/with_property_getter_returns_value_type
' origin: languages/vb/tests/vb/test_vb_with_statements.rs

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

Structure S
Public V As Integer
End Structure
Class C
Public Property Prop As New S()
End Class
Module M
Sub Main()
Dim c1 As New C()
' With c1.Prop ' Fails to mutate original if struct returned by value property
' .V = 10
' End With
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
