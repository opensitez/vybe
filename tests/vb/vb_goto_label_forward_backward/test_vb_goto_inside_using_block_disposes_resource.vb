' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_inside_using_block_disposes_resource
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

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
        Using ms As New MemoryStream()
            ms.WriteByte(1)
            GoTo ExitUsing
        End Using
ExitUsing:
        __Check(CStr("Exited Using Block via GoTo"), "Exited Using Block via GoTo")
    End Sub
End Module
