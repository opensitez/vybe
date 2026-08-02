' vybe-test: vb/vb_modules_namespaces/ns_module_member_access_via_full_qualifier
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

Namespace Services
Public Module MathSvc
Public Function AddOne(v As Integer) As Integer
Return v + 1
End Function
End Module
End Namespace
Module M
Sub Main()
__Check(CStr(Services.MathSvc.AddOne(6)), "7")
End Sub
End Module
