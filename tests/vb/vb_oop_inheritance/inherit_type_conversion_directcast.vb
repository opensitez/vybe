' vybe-test: vb/vb_oop_inheritance/inherit_type_conversion_directcast
' origin: languages/vb/tests/vb/test_vb_oop_inheritance.rs

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
End Class
Class C
Inherits B
End Class
Module M
Sub Main()
Dim b1 As B = New C()
Dim c1 = DirectCast(b1, C)
__Check(CStr(c1 IsNot Nothing), "True")
End Sub
End Module
