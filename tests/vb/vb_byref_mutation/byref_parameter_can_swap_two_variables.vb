' vybe-test: vb/vb_byref_mutation/byref_parameter_can_swap_two_variables
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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
    Sub Swap(ByRef left As Integer, ByRef right As Integer)
        Dim temp As Integer = left
        left = right
        right = temp
    End Sub

    Sub Main()
        Dim a As Integer = 3
        Dim b As Integer = 8
        Swap(a, b)
        __Check(CStr(a), "8")
        __Check(CStr(b), "3")
    End Sub
End Module
