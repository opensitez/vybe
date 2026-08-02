' vybe-test: vb/vb_modules_namespaces/ns_namespace_blocks_share_symbols_across_fragments
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

Namespace Bridge
Public Class Left
Public Value As Integer = 2
End Class
End Namespace
Namespace Bridge
Public Module Builders
Public Function LeftValue() As Integer
Return New Left().Value
End Function
End Module
End Namespace
Module M
Sub Main()
__Check(CStr(Bridge.Builders.LeftValue()), "2")
End Sub
End Module
