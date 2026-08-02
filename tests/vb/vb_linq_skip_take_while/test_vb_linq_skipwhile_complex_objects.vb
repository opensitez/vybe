' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_skipwhile_complex_objects
' origin: languages/vb/tests/vb/test_vb_linq_skip_take_while.rs

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

Imports System.Linq

Class LogEntry
    Public Property Level As String
    Public Sub New(l As String) : Level = l : End Sub
End Class

Module Program
    Sub Main()
        Dim logs = {New LogEntry("INFO"), New LogEntry("INFO"), New LogEntry("ERROR"), New LogEntry("INFO")}
        Dim startingFromError = logs.SkipWhile(Function(l) l.Level = "INFO")
        __Check(CStr(startingFromError.First().Level & "|" & startingFromError.Count()), "ERROR|2")
    End Sub
End Module
