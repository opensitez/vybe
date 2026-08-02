' vybe-test: vb/vb_extension_methods_adv/extension_methods_adv
' origin: languages/vb/tests/vb/test_vb_extension_methods_adv.rs

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

Imports System.Runtime.CompilerServices

Module StringExtensions
    <Extension()>
    Public Function Whisper(ByVal str As String) As String
        Return str.ToLower() & "..."
    End Function
End Module

Module M
    Sub Main()
        Dim msg As String = "HELLO"
        __Check(CStr(msg.Whisper()), "hello...")
    End Sub
End Module
