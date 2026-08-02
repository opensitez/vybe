' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_list_cache_per_type
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class Cache(Of T)
    Public Shared Items As New List(Of T)()
End Class

Module Program
    Sub Main()
        Cache(Of Integer).Items.Add(10)
        Cache(Of Integer).Items.Add(20)
        Cache(Of String).Items.Add("Alpha")

        __Check(CStr(Cache(Of Integer).Items.Count & "|" & Cache(Of String).Items.Count), "2|1")
    End Sub
End Module
