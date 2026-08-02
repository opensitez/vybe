' vybe-test: vb/vb_reflection_property_value_access/test_vb_reflection_get_set_property_value
' origin: languages/vb/tests/vb/test_vb_reflection_property_value_access.rs

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

Imports System.Reflection

Class Configuration
    Public Property ServerHost As String = "127.0.0.1"
End Class

Module Program
    Sub Main()
        Dim cfg As New Configuration()
        Dim t As Type = cfg.GetType()
        Dim prop As PropertyInfo = t.GetProperty("ServerHost")

        __Check(CStr(prop.GetValue(cfg)), "127.0.0.1")
        prop.SetValue(cfg, "192.168.1.1")
        __Check(CStr(cfg.ServerHost), "192.168.1.1")
    End Sub
End Module
