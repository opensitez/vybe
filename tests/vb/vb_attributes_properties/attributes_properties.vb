' vybe-test: vb/vb_attributes_properties/attributes_properties
' origin: languages/vb/tests/vb/test_vb_attributes_properties.rs

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

Class Data
    ' Attributes on properties
    <System.Obsolete("Use NewId instead")>
    Public Property Id As Integer
    
    Public Property NewId As Integer
End Class

Module M
    Sub Main()
        Dim d As New Data()
        d.Id = 10
        __Check(CStr(d.Id), "10")
    End Sub
End Module
