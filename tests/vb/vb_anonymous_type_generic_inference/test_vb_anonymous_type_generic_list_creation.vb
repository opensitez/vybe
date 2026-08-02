' vybe-test: vb/vb_anonymous_type_generic_inference/test_vb_anonymous_type_generic_list_creation
' origin: languages/vb/tests/vb/test_vb_anonymous_type_generic_inference.rs

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

Imports System.Collections.Generic

Module Program
    Private Function CreateList(Of T)(ParamArray items As T()) As List(Of T)
        Return New List(Of T)(items)
    End Function

    Sub Main()
        Dim item1 = New With {.ID = 1}
        Dim item2 = New With {.ID = 2}
        Dim list = CreateList(item1, item2)
        __Check(CStr(list.Count & ":" & list(0).ID & "," & list(1).ID), "2:1,2")
    End Sub
End Module
