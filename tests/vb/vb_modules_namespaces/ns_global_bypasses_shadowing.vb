' vybe-test: vb/vb_modules_namespaces/ns_global_bypasses_shadowing
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
Public Class C
Public V As Integer = 1
End Class
End Namespace
Class N1
Public V As Integer = 2
End Class
Module M
Sub Main()
Dim c As New Global.N1.C()
__Check(CStr(c.V), "1")
End Sub
End Module
