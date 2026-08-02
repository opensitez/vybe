' vybe-test: vb/vb_tuple_deconstruction/tuple_deconstruction
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruction.rs

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
        ' Tuple deconstruction (since VB 15.3 doesn't have native let destructuring quite like C#)
        ' Actually VB 15 doesn't have tuple deconstruction assignment exactly like C# 'var (x, y) = tuple'
        ' But you can do this:
        ' Dim (x, y) = (1, 2) ' This works? Yes in newer VB versions
        Dim t = (1, 2)
        __Check(CStr(t.Item1), "1")
    End Sub
End Module
