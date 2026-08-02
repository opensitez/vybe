' vybe-test: vb/vb_oop_interfaces/interface_multiple_interfaces
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
Sub M1()
End Interface
Interface I2
Sub M2()
End Interface
Class C
Implements I1, I2
Public Sub M1() Implements I1.M1
__Check(CStr("1"), "1")
End Sub
Public Sub M2() Implements I2.M2
__Check(CStr("2"), "2")
End Sub
End Class
Module M
Sub Main()
Dim c1 As New C()
c1.M1()
c1.M2()
End Sub
End Module
