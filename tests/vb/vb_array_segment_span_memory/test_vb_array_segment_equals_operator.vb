' vybe-test: vb/vb_array_segment_span_memory/test_vb_array_segment_equals_operator
' origin: languages/vb/tests/vb/test_vb_array_segment_span_memory.rs

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

Module Program
    Sub Main()
        Dim numbers As Integer() = {10, 20, 30}
        Dim seg1 As New ArraySegment(Of Integer)(numbers, 0, 2)
        Dim seg2 As New ArraySegment(Of Integer)(numbers, 0, 2)
        __Check(CStr(seg1 = seg2), "True")
    End Sub
End Module
