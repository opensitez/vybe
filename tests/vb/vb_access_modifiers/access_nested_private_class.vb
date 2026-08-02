' vybe-test: vb/vb_access_modifiers/access_nested_private_class
' origin: languages/vb/tests/vb/test_vb_access_modifiers.rs

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

Class Outer
Private Class Inner
Public V As Integer = 60
End Class
Public Function GetInnerV() As Integer
Dim i As New Inner()
Return i.V
End Function
End Class
Module M
Sub Main()
Dim o As New Outer()
__Check(CStr(o.GetInnerV()), "60")
End Sub
End Module
