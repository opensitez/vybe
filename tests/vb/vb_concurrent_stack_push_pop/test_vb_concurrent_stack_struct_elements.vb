' vybe-test: vb/vb_concurrent_stack_push_pop/test_vb_concurrent_stack_struct_elements
' origin: languages/vb/tests/vb/test_vb_concurrent_stack_push_pop.rs

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

Imports System.Collections.Concurrent

Structure FrameInfo
    Public FrameId As Integer
    Public Symbol As String
End Structure

Module Program
    Sub Main()
        Dim s As New ConcurrentStack(Of FrameInfo)()
        s.Push(New FrameInfo With {.FrameId = 1, .Symbol = "Main"})
        s.Push(New FrameInfo With {.FrameId = 2, .Symbol = "SubRoutine"})

        Dim info As FrameInfo
        s.TryPop(info)
        __Check(CStr(info.FrameId & ":" & info.Symbol), "2:SubRoutine")
    End Sub
End Module
