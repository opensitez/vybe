' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_static_property_via_instance
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
    Class AppGlobal
        Public Shared Property Counter As Integer = 42
    End Class

    Sub Main()
        Dim obj As Object = New AppGlobal()
        ' Late bound access via instance dispatches to shared property!
        __Check(CStr(CInt(obj.Counter)), "42")
    End Sub
End Module
