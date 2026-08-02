' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_array_type_casting
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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
    Private Function CastArray(Of T)(arr As Array) As T()
        Return CType(arr, T())
    End Function

    Sub Main()
        Dim rawArr As Array = New String() {"A", "B"}
        Dim strArr As String() = CastArray(Of String)(rawArr)
        __Check(CStr(String.Join(",", strArr)), "A,B")
    End Sub
End Module
