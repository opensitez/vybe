' vybe-test: vb/vb_with_statements/with_basic_property_assignment
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

Class C
Public Property V1 As Integer
Public Property V2 As Integer
End Class
Module M
Sub Main()
Dim c1 As New C()
With c1
.V1 = 1
.V2 = 2
End With
__Check(CStr(c1.V1 + c1.V2), "3")
End Sub
End Module
