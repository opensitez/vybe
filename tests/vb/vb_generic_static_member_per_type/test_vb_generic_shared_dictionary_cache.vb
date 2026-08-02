' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_dictionary_cache
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

Class CacheRepository(Of TKey, TValue)
    Public Shared Lookup As New Dictionary(Of TKey, TValue)()
End Class

Module Program
    Sub Main()
        CacheRepository(Of String, Integer).Lookup("Key1") = 100
        CacheRepository(Of Integer, String).Lookup(1) = "One"

        __Check(CStr(CacheRepository(Of String, Integer).Lookup("Key1") & "|" & CacheRepository(Of Integer, String).Lookup(1)), "100|One")
    End Sub
End Module
