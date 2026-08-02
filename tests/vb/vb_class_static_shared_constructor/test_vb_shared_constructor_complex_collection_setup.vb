' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_complex_collection_setup
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class CacheManager
    Public Shared Lookup As New Dictionary(Of String, Integer)()
    Shared Sub New()
        Lookup.Add("Key1", 100)
        Lookup.Add("Key2", 200)
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(CacheManager.Lookup("Key1") + CacheManager.Lookup("Key2")), "300")
    End Sub
End Module
