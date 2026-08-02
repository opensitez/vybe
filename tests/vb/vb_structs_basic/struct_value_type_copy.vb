' vybe-test: vb/vb_structs_basic/struct_value_type_copy
' origin: languages/vb/tests/vb/test_vb_structs_basic.rs

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

Module M
    Sub Main()
        Dim p1 As Point
        p1.X = 10
        p1.Y = 20
        
        ' Assignment creates a copy
        Dim p2 As Point = p1
        p2.X = 50
        
        __Check(CStr(p1.X), "10")
        __Check(CStr(p2.X), "50")
    End Sub
End Module
