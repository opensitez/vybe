' vybe-test: vb/vb_byref_out_advanced/byref_out_parameter
' origin: languages/vb/tests/vb/test_vb_byref_out_advanced.rs

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

Imports System.Runtime.InteropServices

Module M
    ' VB supports <Out> attribute for interop / pseudo-out parameters
    Sub GetValues(ByRef a As Integer, <Out> ByRef b As Integer)
        a = 10
        b = 20
    End Sub

    Sub Main()
        Dim x, y As Integer
        GetValues(x, y)
        __Check(CStr(x.ToString() & " " & y.ToString()), "10 20")
    End Sub
End Module
