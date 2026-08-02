' vybe-test: vb/vb_oop_interfaces/interface_type_conversion_trycast
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
End Interface
Class C
End Class
Module M
Sub Main()
Dim c1 As New C()
Dim i1 = TryCast(c1, I)
__Check(CStr(i1 Is Nothing), "True")
End Sub
End Module
