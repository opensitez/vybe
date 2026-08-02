' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_value_type_struct_boxing
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pt As Object = New Point With {.X = 10, .Y = 20}
        Dim fieldX = GetType(Point).GetField("X")
        fieldX.SetValue(pt, 99)
        Dim unboxed As Point = CType(pt, Point)
        __Check(CStr(unboxed.X), "99")
    End Sub
End Module
