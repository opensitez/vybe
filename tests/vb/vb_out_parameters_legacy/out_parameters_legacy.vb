' vybe-test: vb/vb_out_parameters_legacy/out_parameters_legacy
' origin: languages/vb/tests/vb/test_vb_out_parameters_legacy.rs

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
    ' Legacy way of defining Out parameters in VB
    Sub GetValues(<Out> ByRef a As Integer, <Out> ByRef b As String)
        a = 100
        b = "Data"
    End Sub

    Sub Main()
        Dim a As Integer
        Dim b As String = Nothing
        GetValues(a, b)
        __Check(CStr(a), "100")
        __Check(CStr(b), "Data")
    End Sub
End Module
