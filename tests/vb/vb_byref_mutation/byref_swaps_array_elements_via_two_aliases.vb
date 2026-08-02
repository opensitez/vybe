' vybe-test: vb/vb_byref_mutation/byref_swaps_array_elements_via_two_aliases
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
        Dim saved As Integer = left
        left = right
        right = saved
    End Sub

    Sub Main()
        Dim values() As Integer = New Integer() {8, 1}
        Swap(values(0), values(1))
        __Check(CStr(values(0)), "1")
        __Check(CStr(values(1)), "8")
    End Sub
End Module
