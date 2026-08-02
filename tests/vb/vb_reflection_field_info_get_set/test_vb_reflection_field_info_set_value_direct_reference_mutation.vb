' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_set_value_direct_reference_mutation
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

Class Node
    Public NextNode As Node
    Public Name As String
    Public Sub New(n As String) : Name = n : End Sub
End Class

Module Program
    Sub Main()
        Dim n1 As New Node("N1")
        Dim n2 As New Node("N2")
        Dim field = GetType(Node).GetField("NextNode")
        field.SetValue(n1, n2)
        __Check(CStr(n1.NextNode.Name), "N2")
    End Sub
End Module
