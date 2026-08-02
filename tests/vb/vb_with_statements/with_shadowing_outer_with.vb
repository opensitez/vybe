' vybe-test: vb/vb_with_statements/with_shadowing_outer_with
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

Class C1
Public V As Integer = 1
End Class
Class C2
Public V As Integer = 2
End Class
Module M
Sub Main()
Dim c1 As New C1(), c2 As New C2()
With c1
With c2
__Check(CStr(.V), "2")
End With
End With
End Sub
End Module
