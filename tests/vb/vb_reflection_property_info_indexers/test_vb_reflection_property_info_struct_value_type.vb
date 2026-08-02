' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_struct_value_type
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Structure Point
    Public Property X As Integer
    Public Property Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pt As Object = New Point With {.X = 10, .Y = 20}
        Dim prop = GetType(Point).GetProperty("X")
        prop.SetValue(pt, 50)
        Dim unboxed As Point = CType(pt, Point)
        __Check(CStr(unboxed.X), "50")
    End Sub
End Module
