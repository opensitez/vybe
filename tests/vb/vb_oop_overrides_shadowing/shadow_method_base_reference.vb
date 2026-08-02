' vybe-test: vb/vb_oop_overrides_shadowing/shadow_method_base_reference
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
Public Function M() As String
Return "B"
End Function
End Class
Class C
Inherits B
Public Shadows Function M() As String
Return "C"
End Function
End Class
Module M
Sub Main()
Dim b1 As B = New C()
__Check(CStr(b1.M()), "B")
End Sub
End Module
