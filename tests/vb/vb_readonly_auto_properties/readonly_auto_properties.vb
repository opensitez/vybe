' vybe-test: vb/vb_readonly_auto_properties/readonly_auto_properties
' origin: languages/vb/tests/vb/test_vb_readonly_auto_properties.rs

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
    ' ReadOnly auto-property can be initialized inline
    Public ReadOnly Property Id As Integer = 42
    Public ReadOnly Property Name As String
    
    Public Sub New(name As String)
        ' Or initialized in the constructor
        Me.Name = name
    End Sub
End Class

Module M
    Sub Main()
        Dim i As New Item("Test")
        __Check(CStr(i.Id), "42")
        __Check(CStr(i.Name), "Test")
    End Sub
End Module
