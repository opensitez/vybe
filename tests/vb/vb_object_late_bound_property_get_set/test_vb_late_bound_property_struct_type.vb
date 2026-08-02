' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_property_struct_type
' origin: languages/vb/tests/vb/test_vb_object_late_bound_property_get_set.rs

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

Module Program
    Structure Size2D
        Public Width As Integer
        Public Height As Integer
    End Structure

    Class Frame
        Public Property Bounds As Size2D
    End Class

    Sub Main()
        Dim obj As Object = New Frame()
        obj.Bounds = New Size2D With {.Width = 1920, .Height = 1080}
        Dim sz As Size2D = CType(obj.Bounds, Size2D)
        __Check(CStr(sz.Width & "x" & sz.Height), "1920x1080")
    End Sub
End Module
