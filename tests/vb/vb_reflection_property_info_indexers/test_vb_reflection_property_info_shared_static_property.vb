' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_shared_static_property
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Class SystemState
    Public Shared Property AppName As String = "VybeApp"
End Class

Module Program
    Sub Main()
        Dim prop = GetType(SystemState).GetProperty("AppName")
        __Check(CStr(prop.GetValue(Nothing)), "VybeApp")
        prop.SetValue(Nothing, "NewVybeApp")
        __Check(CStr(SystemState.AppName), "NewVybeApp")
    End Sub
End Module
