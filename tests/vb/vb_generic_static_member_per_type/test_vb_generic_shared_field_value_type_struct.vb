' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_value_type_struct
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Structure GenericStructHolder(Of T)
    Public Shared DefaultValue As T
End Structure

Module Program
    Sub Main()
        GenericStructHolder(Of Integer).DefaultValue = 99
        GenericStructHolder(Of String).DefaultValue = "DefaultText"

        __Check(CStr(GenericStructHolder(Of Integer).DefaultValue & "|" & GenericStructHolder(Of String).DefaultValue), "99|DefaultText")
    End Sub
End Module
