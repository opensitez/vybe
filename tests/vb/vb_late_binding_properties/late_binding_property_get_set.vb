' vybe-test: vb/vb_late_binding_properties/late_binding_property_get_set
' origin: languages/vb/tests/vb/test_vb_late_binding_properties.rs

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

Class DataModel
    Private _value As String
    Public Property Value As String
        Get
            Return _value
        End Get
        Set(v As String)
            _value = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim model As Object = New DataModel()
        
        ' Late bound property setter
        model.Value = "Dynamic Property"
        
        ' Late bound property getter
        __Check(CStr(model.Value), "Dynamic Property")
    End Sub
End Module
