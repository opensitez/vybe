' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_set_length_truncate
' origin: languages/vb/tests/vb/test_vb_memory_stream_capacity_expansion.rs

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

Imports System.IO

Module Program
    Sub Main()
        Using ms As New MemoryStream()
            ms.Write({1, 2, 3, 4, 5}, 0, 5)
            ms.SetLength(3)
            __Check(CStr(ms.Length & "|" & String.Join(",", ms.ToArray())), "3|1,2,3")
        End Using
    End Sub
End Module
