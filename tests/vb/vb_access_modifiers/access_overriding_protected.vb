' vybe-test: vb/vb_access_modifiers/access_overriding_protected
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

Class B
Protected Overridable Function M1() As String
Return "B"
End Function
End Class
Class C
Inherits B
Protected Overrides Function M1() As String
Return "C"
End Function
Public Function CallM1() As String
Return M1()
End Function
End Class
Module M
Sub Main()
Dim c1 As New C()
__Check(CStr(c1.CallM1()), "C")
End Sub
End Module
