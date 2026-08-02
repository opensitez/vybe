' vybe-test: vb/vb_oop_interfaces/interface_overloading_methods
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
Function M(v As Integer) As String
Function M(v As String) As String
End Interface
Class C
Implements I
Public Function M(v As Integer) As String Implements I.M
Return "Int"
End Function
Public Function M(v As String) As String Implements I.M
Return "Str"
End Function
End Class
Module M
Sub Main()
Dim c1 As I = New C()
__Check(CStr(c1.M(5)), "Int")
End Sub
End Module
