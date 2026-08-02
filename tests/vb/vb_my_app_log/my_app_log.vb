' vybe-test: vb/vb_my_app_log/my_app_log
' origin: languages/vb/tests/vb/test_vb_my_app_log.rs

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
    ' Simulate My.Application.Log
    Class LogClass
        Public Sub WriteEntry(msg As String)
            __Check(CStr("Log: " & msg), "Log: Started")
        End Sub
    End Class
    
    Class AppClass
        Public Log As New LogClass()
    End Class
    
    Dim Application As New AppClass()
    
    Sub Main()
        Application.Log.WriteEntry("Started")
    End Sub
End Module
