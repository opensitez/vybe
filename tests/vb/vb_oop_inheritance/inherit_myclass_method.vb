' vybe-test: vb/vb_oop_inheritance/inherit_myclass_method
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
Public Overridable Function GetV() As Integer
Return 1
End Function
Public Function CallMyClass() As Integer
Return MyClass.GetV()
End Function
End Class
Class C
Inherits B
Public Overrides Function GetV() As Integer
Return 2
End Function
End Class
Module M
Sub Main()
Dim c1 As New C()
__Check(CStr(c1.CallMyClass()), "1")
End Sub
End Module
