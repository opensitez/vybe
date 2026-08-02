' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_delegate_invocation
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

Delegate Function Transform(Of T, R)(item As T) As R

Module Program
    Sub Main()
        Dim stringLen As Transform(Of String, Integer) = Function(s) s.Length
        __Check(CStr(stringLen("VisualBasic")), "11")
    End Sub
End Module
