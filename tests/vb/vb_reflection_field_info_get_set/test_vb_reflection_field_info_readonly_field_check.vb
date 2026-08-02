' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_readonly_field_check
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

Class ReadOnlyContainer
    Public ReadOnly ID As Integer = 42
    Public Sub New(idVal As Integer) : ID = idVal : End Sub
End Class

Module Program
    Sub Main()
        Dim field = GetType(ReadOnlyContainer).GetField("ID")
        __Check(CStr(field.IsInitOnly), "True")
    End Sub
End Module
