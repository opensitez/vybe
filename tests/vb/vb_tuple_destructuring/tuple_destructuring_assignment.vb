' vybe-test: vb/vb_tuple_destructuring/tuple_destructuring_assignment
' origin: languages/vb/tests/vb/test_vb_tuple_destructuring.rs

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
    Function GetInfo() As (Name As String, Age As Integer)
        Return ("John", 30)
    End Function

    Sub Main()
        ' VB.NET does not natively support deconstruction syntax (Dim (name, age) = GetInfo())
        ' Wait, it doesn't? C# has deconstruction, but VB.NET does not have direct tuple deconstruction assignment syntax.
        ' Let's just use the tuple literal syntax and element access.
        Dim t = GetInfo()
        __Check(CStr(t.Name), "John")
        __Check(CStr(t.Age), "30")
        
        ' We can assign tuples to tuples
        Dim t2 As (String, Integer) = t
        __Check(CStr(t2.Item1), "John")
    End Sub
End Module
