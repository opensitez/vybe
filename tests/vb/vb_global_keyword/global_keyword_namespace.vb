' vybe-test: vb/vb_global_keyword/global_keyword_namespace
' origin: languages/vb/tests/vb/test_vb_global_keyword.rs

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

Namespace System
    Class Console
        Public Shared Sub WriteLine(s As String)
            ' Shadowing the real System.Console
        End Sub
    End Class
End Namespace

Module M
    Sub Main()
        ' Using Global allows escaping the local namespace shadowing to hit the root
        Global.System.__Check(CStr("Hit Root"), "Hit Root")
    End Sub
End Module
