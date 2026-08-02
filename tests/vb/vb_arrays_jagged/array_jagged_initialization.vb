' vybe-test: vb/vb_arrays_jagged/array_jagged_initialization
' origin: languages/vb/tests/vb/test_vb_arrays_jagged.rs

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
        Dim jagged As Integer()() = {
            New Integer() {10, 20},
            New Integer() {30, 40, 50}
        }
        
        __Check(CStr(jagged(0)(1)), "20")
        __Check(CStr(jagged(1)(0)), "30")
    End Sub
End Module
