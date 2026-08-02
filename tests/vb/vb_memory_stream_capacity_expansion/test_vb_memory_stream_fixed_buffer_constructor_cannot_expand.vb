' vybe-test: vb/vb_memory_stream_capacity_expansion/test_vb_memory_stream_fixed_buffer_constructor_cannot_expand
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

Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim fixedBuffer As Byte() = New Byte(4) {}
        Using ms As New MemoryStream(fixedBuffer)
            Try
                ms.SetLength(10)
            Catch ex As NotSupportedException
                __Check(CStr("NotSupportedException Caught on Fixed MemoryStream"), "NotSupportedException Caught on Fixed MemoryStream")
            End Try
        End Using
    End Sub
End Module
