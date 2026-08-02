' vybe-test: vb/vb_property_backing_field_custom/test_vb_property_validation_in_setter
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

Class AgeTracker
    Private _age As Integer

    Public Property Age As Integer
        Get
            Return _age
        End Get
        Set(value As Integer)
            If value < 0 Then
                _age = 0
            Else
                _age = value
            End If
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim t As New AgeTracker()
        t.Age = -5
        __Check(CStr(t.Age), "0")
        t.Age = 25
        __Check(CStr(t.Age), "25")
    End Sub
End Module
