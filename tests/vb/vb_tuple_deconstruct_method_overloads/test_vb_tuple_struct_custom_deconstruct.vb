' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_struct_custom_deconstruct
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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

Structure Dimensions
    Public Width As Integer
    Public Height As Integer
    Public Sub New(w As Integer, h As Integer) : Width = w : Height = h : End Sub
    Public Sub Deconstruct(ByRef w As Integer, ByRef h As Integer)
        w = Width : h = Height
    End Sub
End Structure

Module Program
    Sub Main()
        Dim d As New Dimensions(1920, 1080)
        Dim w As Integer = 0, h As Integer = 0
        d.Deconstruct(w, h)
        __Check(CStr(w & "x" & h), "1920x1080")
    End Sub
End Module
