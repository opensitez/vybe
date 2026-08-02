' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_field_custom_struct_value
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

Structure ConfigData
    Public ID As Integer
    Public Name As String
End Structure

Class SystemConfig(Of T)
    Public Shared Config As ConfigData
End Class

Module Program
    Sub Main()
        SystemConfig(Of Integer).Config = New ConfigData With {.ID = 1, .Name = "IntCfg"}
        SystemConfig(Of String).Config = New ConfigData With {.ID = 2, .Name = "StrCfg"}

        __Check(CStr(SystemConfig(Of Integer).Config.Name & "|" & SystemConfig(Of String).Config.Name), "IntCfg|StrCfg")
    End Sub
End Module
