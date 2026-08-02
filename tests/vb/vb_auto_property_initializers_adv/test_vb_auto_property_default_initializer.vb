' vybe-test: vb/vb_auto_property_initializers_adv/test_vb_auto_property_default_initializer
' origin: languages/vb/tests/vb/test_vb_auto_property_initializers_adv.rs

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

Class Config
    Public Property Port As Integer = 8080
    Public Property Host As String = "localhost"
    Public Property IsEnabled As Boolean = True
End Class

Module Program
    Sub Main()
        Dim cfg As New Config()
        __Check(CStr(cfg.Host & ":" & cfg.Port & ":" & cfg.IsEnabled), "localhost:8080:True")
    End Sub
End Module
