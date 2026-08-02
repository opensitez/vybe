' vybe-test: vb/vb_spec_namespaces_modules/namespace_spec_namespace_can_nest_three_levels_deep
' origin: languages/vb/tests/vb/test_vb_spec_namespaces_modules.rs

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

Namespace One
    Namespace Two
        Namespace Three
            Public Class Box
                Public Function Name() As String
                    Return "deep"
                End Function
            End Class
        End Namespace
    End Namespace
End Namespace
Module M
    Sub Main()
        __Check(CStr((New One.Two.Three.Box()).Name()), "deep")
    End Sub
End Module
