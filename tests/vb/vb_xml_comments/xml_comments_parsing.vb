' vybe-test: vb/vb_xml_comments/xml_comments_parsing
' origin: languages/vb/tests/vb/test_vb_xml_comments.rs

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
    ''' <summary>
    ''' This is a summary.
    ''' </summary>
    ''' <param name="msg">The message to print.</param>
    Sub Main()
        __Check(CStr("Parsed OK"), "Parsed OK")
    End Sub
End Module
