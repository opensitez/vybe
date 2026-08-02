' vybe-test: vb/vb_arrays_jagged/array_jagged_basic
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
        Dim jagged(2)() As Integer
        
        jagged(0) = New Integer() {1, 2}
        jagged(1) = New Integer() {3, 4, 5}
        jagged(2) = New Integer() {6}
        
        __Check(CStr(jagged(1)(2)), "5")
        __Check(CStr(jagged.Length), "3")
    End Sub
End Module
