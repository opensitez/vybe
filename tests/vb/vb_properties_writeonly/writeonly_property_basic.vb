' vybe-test: vb/vb_properties_writeonly/writeonly_property_basic
' origin: languages/vb/tests/vb/test_vb_properties_writeonly.rs

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

Class Logger
    Private _lastMessage As String
    
    Public WriteOnly Property Message As String
        Set(value As String)
            _lastMessage = value
            __Check(CStr("Logged: " & _lastMessage), "Logged: System started")
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim log As New Logger()
        log.Message = "System started"
    End Sub
End Module
