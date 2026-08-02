' vybe-test: vb/vb_global_namespace/global_namespace_access
' origin: languages/vb/tests/vb/test_vb_global_namespace.rs

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

Namespace MyProject.Utils
    Class Logger
        Public Sub Log(msg As String)
            __Check(CStr("Utils Logger: " & msg), "Global Logger: Hello")
        End Sub
    End Class
End Namespace

Class Logger
    Public Sub Log(msg As String)
        __Check(CStr("Global Logger: " & msg), "Utils Logger: Hello")
    End Sub
End Class

Module M
    Sub Main()
        ' Accessing global namespace using the Global keyword
        Dim gLog As New Global.Logger()
        gLog.Log("Hello")
        
        Dim uLog As New Global.MyProject.Utils.Logger()
        uLog.Log("Hello")
    End Sub
End Module
