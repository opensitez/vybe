' vybe-test: vb/vb_concurrent_dictionary_try_add/test_vb_concurrent_dictionary_struct_key_and_value
' origin: languages/vb/tests/vb/test_vb_concurrent_dictionary_try_add.rs

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

Structure Point2D
    Public X, Y As Integer
End Structure

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Point2D, String)()
        Dim p As New Point2D With {.X = 1, .Y = 2}
        dict.TryAdd(p, "Point(1,2)")
        __Check(CStr(dict(p)), "Point(1,2)")
    End Sub
End Module
