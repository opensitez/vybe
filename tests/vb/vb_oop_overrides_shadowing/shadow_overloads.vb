' vybe-test: vb/vb_oop_overrides_shadowing/shadow_overloads
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
Public Sub M(v As Integer)
__Check(CStr("B"), "B")
End Sub
End Class
Class C
Inherits B
Public Overloads Sub M(s As String)
__Check(CStr("C"), "C")
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
c1.M(1)
c1.M("A")
End Sub
End Module
