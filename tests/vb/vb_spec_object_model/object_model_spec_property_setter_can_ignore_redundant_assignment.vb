' vybe-test: vb/vb_spec_object_model/object_model_spec_property_setter_can_ignore_redundant_assignment
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

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

Class Meter
    Private _value As Integer
    Public Property Value As Integer
        Get
            Return _value
        End Get
        Set(value As Integer)
            If _value <> value Then
                _value = value
            End If
        End Set
    End Property
End Class
Module M
    Sub Main()
        Dim meter As New Meter()
        meter.Value = 3
        meter.Value = 3
        __Check(CStr(meter.Value), "3")
    End Sub
End Module
