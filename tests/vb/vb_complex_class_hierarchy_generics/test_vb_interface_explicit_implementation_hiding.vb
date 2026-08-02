' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_interface_explicit_implementation_hiding
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Interface ISecret
    Sub Reveal()
End Interface

Class Vault
    Implements ISecret
    Private Sub RevealSecret() Implements ISecret.Reveal
        __Check(CStr("Secret Revealed"), "Secret Revealed")
    End Sub
End Class

Module Program
    Sub Main()
        Dim v As New Vault()
        Dim s As ISecret = v
        s.Reveal()
    End Sub
End Module
