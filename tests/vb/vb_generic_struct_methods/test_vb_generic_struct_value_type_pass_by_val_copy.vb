' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_value_type_pass_by_val_copy
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure MutableBox(Of T)
    Public Item As T
    Public Sub New(i As T)
        Item = i
    End Sub
End Structure

Module Program
    Private Sub ModifyBox(b As MutableBox(Of Integer))
        b.Item = 99
    End Sub

    Sub Main()
        Dim b As New MutableBox(Of Integer)(10)
        ModifyBox(b)
        __Check(CStr(b.Item), "10")
    End Sub
End Module
