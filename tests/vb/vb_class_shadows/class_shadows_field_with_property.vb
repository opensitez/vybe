' vybe-test: vb/vb_class_shadows/class_shadows_field_with_property
' origin: languages/vb/tests/vb/test_vb_class_shadows.rs

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

Class BaseData
    Public Value As String = "BaseValue"
End Class

Class DerivedData
    Inherits BaseData
    
    Private _val As String = "DerivedValue"
    Public Shadows Property Value As String
        Get
            Return _val
        End Get
        Set(v As String)
            _val = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim d As New DerivedData()
        __Check(CStr(d.Value), "DerivedValue")
        
        Dim b As BaseData = d
        __Check(CStr(b.Value), "BaseValue")
    End Sub
End Module
