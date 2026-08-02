' vybe-test: vb/vb_anonymous_types/anonymous_type_basic
' origin: languages/vb/tests/vb/test_vb_anonymous_types.rs

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
        ' Creates an instance of an anonymous type
        Dim person = New With { .Name = "Alice", .Age = 30 }
        
        __Check(CStr(person.Name), "Alice")
        __Check(CStr(person.Age), "30")
        
        ' In VB, properties of anonymous types without 'Key' modifier are mutable
        person.Age = 31
        __Check(CStr(person.Age), "31")
    End Sub
End Module
