' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_class_deeply_nested_three_levels
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class Level1
    Public Class Level2
        Public Class Level3
            Public Shared Function Hello() As String
                Return "Level 3 Hello"
            End Function
        End Class
    End Class
End Class

Module Program
    Sub Main()
        __Check(CStr(Level1.Level2.Level3.Hello()), "Level 3 Hello")
    End Sub
End Module
