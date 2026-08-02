' vybe-test: vb/vb_oop_classes_constructors/class_shadowing_constructor
' origin: languages/vb/tests/vb/test_vb_oop_classes_constructors.rs

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
Public Sub New(v As Integer)
End Sub
End Class
Class C
Inherits B
Public Sub New()
MyBase.New(1)
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
__Check(CStr(c1 IsNot Nothing), "True")
End Sub
End Module
