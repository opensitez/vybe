' vybe-test: vb/vb_my_settings_sim/my_settings_sim
' origin: languages/vb/tests/vb/test_vb_my_settings_sim.rs

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

Module M
    ' Simulate My.Settings functionality
    Class SettingsClass
        Public DefaultProperty As String = "Value"
    End Class
    
    Dim Settings As New SettingsClass()
    
    Sub Main()
        __Check(CStr(Settings.DefaultProperty), "Value")
    End Sub
End Module
