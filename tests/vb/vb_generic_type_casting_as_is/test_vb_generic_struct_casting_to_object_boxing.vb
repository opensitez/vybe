' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_struct_casting_to_object_boxing
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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
    Private Function BoxValue(Of T)(val As T) As Object
        Return CObj(val)
    End Function

    Sub Main()
        Dim p As New Point With {.X = 1, .Y = 2}
        Dim boxed = BoxValue(p)
        __Check(CStr(TypeOf boxed Is Point), "True")
    End Sub
End Module
