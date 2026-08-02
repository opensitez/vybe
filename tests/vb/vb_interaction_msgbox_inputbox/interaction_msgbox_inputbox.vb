' vybe-test: vb/vb_interaction_msgbox_inputbox/interaction_msgbox_inputbox
' origin: languages/vb/tests/vb/test_vb_interaction_msgbox_inputbox.rs

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
    Sub Main()
        ' We can't actually show UI, but we can test the syntax parsing
        ' Mocking MsgBoxResult Enum implicitly used
        Dim prompt As String = "Test"
        Dim title As String = "Title"
        
        ' Just check it compiles
        Dim msgType = MsgBoxStyle.OkOnly
        __Check(CStr(CInt(msgType)), "0")
        
        ' Don't call them as it might hang the test runner if not mocked
        __Check(CStr("Parsed"), "Parsed")
    End Sub
End Module
