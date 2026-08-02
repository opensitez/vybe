' vybe-test: vb/vb_anonymous_types_collections/anonymous_types_collections
' origin: languages/vb/tests/vb/test_vb_anonymous_types_collections.rs

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

Module M
    Sub Main()
        ' Array of anonymous types
        Dim arr = {
            New With { Key .Id = 1, .Name = "Alice" },
            New With { Key .Id = 2, .Name = "Bob" }
        }
        
        __Check(CStr(arr(0).Name), "Alice")
        __Check(CStr(arr.Length), "2")
        
        ' Collection initializer with anonymous types
        Dim list As New System.Collections.Generic.List(Of Object) From {
            New With { .Value = 10 },
            New With { .Value = 20 }
        }
        __Check(CStr(list.Count), "2")
    End Sub
End Module
