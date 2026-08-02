' vybe-test: vb/vb_oop_overrides_shadowing/shadow_nested_type
' origin: languages/vb/tests/vb/test_vb_oop_overrides_shadowing.rs

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

Class B
Class N
Public V As Integer = 1
End Class
End Class
Class C
Inherits B
Shadows Class N
Public V As Integer = 2
End Class
End Class
Module M
Sub Main()
Dim n1 As New C.N()
__Check(CStr(n1.V), "2")
End Sub
End Module
