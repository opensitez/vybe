' vybe-test: vb/vb_oop_interfaces/interface_implementation_with_byref
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

Interface I
Sub Mutate(ByRef v As Integer)
End Interface
Class C
Implements I
Public Sub Mutate(ByRef v As Integer) Implements I.Mutate
v = 10
End Sub
End Class
Module M
Sub Main()
Dim c1 As I = New C()
Dim x = 1
c1.Mutate(x)
__Check(CStr(x), "10")
End Sub
End Module
