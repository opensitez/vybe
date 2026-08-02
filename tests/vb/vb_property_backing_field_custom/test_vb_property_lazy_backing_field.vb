' vybe-test: vb/vb_property_backing_field_custom/test_vb_property_lazy_backing_field
' origin: languages/vb/tests/vb/test_vb_property_backing_field_custom.rs

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

Class LazyData
    Private _data As String = Nothing

    Public ReadOnly Property Data As String
        Get
            If _data Is Nothing Then
                _data = "ComputedValue"
            End If
            Return _data
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim d As New LazyData()
        __Check(CStr(d.Data), "ComputedValue")
    End Sub
End Module
