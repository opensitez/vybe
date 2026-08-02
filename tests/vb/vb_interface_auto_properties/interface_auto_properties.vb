' vybe-test: vb/vb_interface_auto_properties/interface_auto_properties
' origin: languages/vb/tests/vb/test_vb_interface_auto_properties.rs

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

Interface IData
    ' Properties in interfaces don't need Get/Set explicitly if they are auto-implemented style
    Property Value As Integer
End Interface

Class Data
    Implements IData
    
    Public Property Value As Integer Implements IData.Value
End Class

Module M
    Sub Main()
        Dim d As IData = New Data()
        d.Value = 100
        __Check(CStr(d.Value), "100")
    End Sub
End Module
