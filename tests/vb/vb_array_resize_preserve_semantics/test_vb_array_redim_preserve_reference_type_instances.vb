' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_reference_type_instances
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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

Class Item
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim items(0) As Item
        items(0) = New Item("First")
        ReDim Preserve items(1)
        __Check(CStr(items(0).Name), "First")
        __Check(CStr(items(1) Is Nothing), "True")
    End Sub
End Module
