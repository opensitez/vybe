' vybe-test: vb/vb_late_binding_advanced/late_binding_property_access
' origin: languages/vb/tests/vb/test_vb_late_binding_advanced.rs

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

Class Item
    Public Property Name As String
End Class

Module M
    Sub Main()
        ' With Option Strict Off (default), Object variables can use late binding
        Dim obj As Object = New Item() With { .Name = "TestItem" }
        
        ' Late-bound property access
        __Check(CStr(obj.Name), "TestItem")
    End Sub
End Module
