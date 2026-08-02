' vybe-test: vb/vb_property_default_values/property_default_values
' origin: languages/vb/tests/vb/test_vb_property_default_values.rs

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

Class Configuration
    ' Default values for auto-implemented properties
    Public Property MaxItems As Integer = 100
    Public Property Description As String = "Default Settings"
    Public Property IsEnabled As Boolean = True
End Class

Module M
    Sub Main()
        Dim config As New Configuration()
        __Check(CStr(config.MaxItems), "100")
        __Check(CStr(config.Description), "Default Settings")
        __Check(CStr(config.IsEnabled), "True")
        
        config.MaxItems = 50
        __Check(CStr(config.MaxItems), "50")
    End Sub
End Module
