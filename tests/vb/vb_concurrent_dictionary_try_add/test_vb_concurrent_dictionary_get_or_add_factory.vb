' vybe-test: vb/vb_concurrent_dictionary_try_add/test_vb_concurrent_dictionary_get_or_add_factory
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

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of String, String)()
        Dim s1 = dict.GetOrAdd("User1", Function(k) "NewUser_" & k)
        Dim s2 = dict.GetOrAdd("User1", Function(k) "OtherUser")
        __Check(CStr(s1 & "|" & s2), "NewUser_User1|NewUser_User1")
    End Sub
End Module
