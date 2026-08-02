' vybe-test: vb/vb_array_segment_span_memory/test_vb_array_segment_empty_property
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
        Dim emptySeg As ArraySegment(Of Integer) = ArraySegment(Of Integer).Empty
        __Check(CStr(emptySeg.Count), "0")
    End Sub
End Module
