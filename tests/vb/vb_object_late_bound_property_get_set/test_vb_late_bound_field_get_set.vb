' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_field_get_set
' origin: languages/vb/tests/vb/test_vb_object_late_bound_property_get_set.rs

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

Module Program
    Class Settings
        Public AppName As String
    End Class

    Sub Main()
        Dim obj As Object = New Settings()
        obj.AppName = "VybeEngine"
        Dim name As String = CStr(obj.AppName)
        __Check(CStr(name), "VybeEngine")
    End Sub
End Module
