' vybe-test: vb/vb_class_mustinherit/class_mustoverride_properties
' origin: languages/vb/tests/vb/test_vb_class_mustinherit.rs

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

MustInherit Class Vehicle
    Public MustOverride ReadOnly Property Wheels() As Integer
End Class

Class Tricycle
    Inherits Vehicle
    
    Public Overrides ReadOnly Property Wheels() As Integer
        Get
            Return 3
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim t As Vehicle = New Tricycle()
        __Check(CStr(t.Wheels), "3")
    End Sub
End Module
