' vybe-test: vb/vb_oop_interfaces/interface_as_property_type
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
Function M() As Integer
End Interface
Class C
Implements I
Public Function M() As Integer Implements I.M
Return 42
End Function
End Class
Class Container
Public Property Obj As I
End Class
Module M
Sub Main()
Dim cont As New Container()
cont.Obj = New C()
__Check(CStr(cont.Obj.M()), "42")
End Sub
End Module
