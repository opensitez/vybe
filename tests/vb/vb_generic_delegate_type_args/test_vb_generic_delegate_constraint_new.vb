' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_constraint_new
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Delegate Function Creator(Of T As New)() As T

Class Item
    Public Tag As String = "CreatedItem"
End Class

Module Program
    Sub Main()
        Dim create As Creator(Of Item) = Function() New Item()
        Dim item = create()
        __Check(CStr(item.Tag), "CreatedItem")
    End Sub
End Module
