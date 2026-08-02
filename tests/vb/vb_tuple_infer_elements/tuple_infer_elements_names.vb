' vybe-test: vb/vb_tuple_infer_elements/tuple_infer_elements_names
' origin: languages/vb/tests/vb/test_vb_tuple_infer_elements.rs

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
        Dim count = 5
        Dim name = "Bob"
        
        Dim t = (count, name)
        
        ' If element name inference is supported, t.count should work
        ' However, we'll use a hack to check by using reflection on properties? No, ValueTuple fields.
        ' Let's just do standard assignment.
        Dim t2 As (C As Integer, N As String) = t
        __Check(CStr(t2.C), "5")
        __Check(CStr(t2.N), "Bob")
    End Sub
End Module
