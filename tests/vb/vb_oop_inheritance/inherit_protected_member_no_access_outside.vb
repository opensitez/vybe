' vybe-test: vb/vb_oop_inheritance/inherit_protected_member_no_access_outside
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
Protected V As Integer = 42
End Class
Module M
Sub Main()
Dim b1 As New B()
' b1.V ' Fails
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
