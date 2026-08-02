' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_public_instance_get_set
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

Class DataBox
    Public Tag As String = "Initial"
End Class

Module Program
    Sub Main()
        Dim box As New DataBox()
        Dim field = GetType(DataBox).GetField("Tag")
        __Check(CStr(field.GetValue(box)), "Initial")
        field.SetValue(box, "UpdatedTag")
        __Check(CStr(box.Tag), "UpdatedTag")
    End Sub
End Module
