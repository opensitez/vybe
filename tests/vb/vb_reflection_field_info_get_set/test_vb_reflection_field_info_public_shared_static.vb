' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_public_shared_static
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

Class GlobalConfig
    Public Shared Version As String = "1.0.0"
End Class

Module Program
    Sub Main()
        Dim field = GetType(GlobalConfig).GetField("Version")
        __Check(CStr(field.GetValue(Nothing)), "1.0.0")
        field.SetValue(Nothing, "2.0.0")
        __Check(CStr(GlobalConfig.Version), "2.0.0")
    End Sub
End Module
