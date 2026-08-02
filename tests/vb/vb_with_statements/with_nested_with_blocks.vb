' vybe-test: vb/vb_with_statements/with_nested_with_blocks
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

Class Inner
Public V As Integer = 5
End Class
Class Outer
Public Prop As New Inner()
End Class
Module M
Sub Main()
Dim o As New Outer()
With o
With .Prop
__Check(CStr(.V), "5")
End With
End With
End Sub
End Module
