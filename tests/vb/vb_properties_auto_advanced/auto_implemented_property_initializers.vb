' vybe-test: vb/vb_properties_auto_advanced/auto_implemented_property_initializers
' origin: languages/vb/tests/vb/test_vb_properties_auto_advanced.rs

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
    ' Auto-property with initializer
    Public Property Name As String = "Unknown"
    ' Auto-property with object initializer syntax for collection
    Public Property Tags As New System.Collections.Generic.List(Of String) From {"New"}
End Class

Module M
    Sub Main()
        Dim i As New Item()
        __Check(CStr(i.Name), "Unknown")
        __Check(CStr(i.Tags(0)), "New")
        
        i.Name = "Known"
        __Check(CStr(i.Name), "Known")
    End Sub
End Module
