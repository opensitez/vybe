' vybe-test: vb/vb_system_threading/system_threading_interlocked
' origin: languages/vb/tests/vb/test_vb_system_threading.rs

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

Imports System.Threading

Module M
    Sub Main()
        Dim count As Integer = 0
        Interlocked.Increment(count)
        __Check(CStr(count), "1")
        
        Interlocked.Add(count, 5)
        __Check(CStr(count), "6")
    End Sub
End Module
