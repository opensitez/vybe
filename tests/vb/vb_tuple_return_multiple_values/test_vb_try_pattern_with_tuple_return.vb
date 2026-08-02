' vybe-test: vb/vb_tuple_return_multiple_values/test_vb_try_pattern_with_tuple_return
' origin: languages/vb/tests/vb/test_vb_tuple_return_multiple_values.rs

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

Module Program
    Function TryParseInt(input As String) As (Success As Boolean, Value As Integer)
        Dim result As Integer
        If Integer.TryParse(input, result) Then
            Return (True, result)
        End If
        Return (False, 0)
    End Function

    Sub Main()
        Dim r1 = TryParseInt("123")
        Dim r2 = TryParseInt("abc")
        __Check(CStr(r1.Success & ":" & r1.Value), "True:123")
        __Check(CStr(r2.Success & ":" & r2.Value), "False:0")
    End Sub
End Module
