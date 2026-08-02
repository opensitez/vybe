' vybe-test: vb/vb_modules_namespaces/ns_multiple_blocks_same_name
' origin: languages/vb/tests/vb/test_vb_modules_namespaces.rs

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

Namespace N1
Public Class C1
Public V1 As Integer = 1
End Class
End Namespace
Namespace N1
Public Class C2
Public V2 As Integer = 2
End Class
End Namespace
Module M
Sub Main()
Dim c1 As New N1.C1()
Dim c2 As New N1.C2()
__Check(CStr(c1.V1 + c2.V2), "3")
End Sub
End Module
