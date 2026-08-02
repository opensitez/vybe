' vybe-test: vb/vb_modules_namespaces/ns_namespace_alias_to_type_reference
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

Namespace Data
Public Class Packet
Public Status As String = "ready"
End Class
End Namespace
Imports PacketAlias = Data.Packet
Module M
Sub Main()
Dim p As New PacketAlias()
__Check(CStr(p.Status), "ready")
End Sub
End Module
