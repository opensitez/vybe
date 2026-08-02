' vybe-test: vb/vb_oop_interfaces/interface_implement_multiple_methods_with_one
' origin: languages/vb/tests/vb/test_vb_oop_interfaces.rs

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

Interface I1
Sub M()
End Interface
Interface I2
Sub M()
End Interface
Class C
Implements I1, I2
Public Sub M() Implements I1.M, I2.M
__Check(CStr("C"), "C")
End Sub
End Class
Module M
Sub Main()
Dim c1 As I2 = New C()
c1.M()
End Sub
End Module
