' vybe-test: vb/vb_array_empty_and_null_bounds/test_vb_array_clone_reference_types_shallow
' origin: languages/vb/tests/vb/test_vb_array_empty_and_null_bounds.rs

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

Class Container
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim orig As Container() = {New Container("A")}
        Dim cloned As Container() = CType(orig.Clone(), Container())
        cloned(0).Tag = "Modified"
        __Check(CStr(orig(0).Tag), "Modified")
    End Sub
End Module
