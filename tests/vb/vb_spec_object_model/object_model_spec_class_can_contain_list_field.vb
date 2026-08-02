' vybe-test: vb/vb_spec_object_model/object_model_spec_class_can_contain_list_field
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

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

Class Bag
    Public Items As New List(Of Integer)
End Class
Module M
    Sub Main()
        Dim bag As New Bag()
        bag.Items.Add(9)
        __Check(CStr(bag.Items.Count), "1")
    End Sub
End Module
