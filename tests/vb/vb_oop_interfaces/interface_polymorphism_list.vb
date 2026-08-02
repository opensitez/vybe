' vybe-test: vb/vb_oop_interfaces/interface_polymorphism_list
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
Function M() As String
End Interface
Class A
Implements I
Public Function M() As String Implements I.M
Return "A"
End Function
End Class
Class B
Implements I
Public Function M() As String Implements I.M
Return "B"
End Function
End Class
Module M
Sub Main()
Dim arr() As I = {New A(), New B()}
__Check(CStr(arr(0).M() & arr(1).M()), "AB")
End Sub
End Module
