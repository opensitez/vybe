' vybe-test: vb/vb_string_padleft_padright_formatting/test_vb_string_padright_table_column_formatting
' origin: languages/vb/tests/vb/test_vb_string_padleft_padright_formatting.rs

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
    Sub Main()
        Dim col1 As String = "Name"
        Dim col2 As String = "Age"
        Dim row1Name As String = "Alice"
        Dim row1Age As String = "30"

        __Check(CStr(col1.PadRight(10) & "|" & col2.PadLeft(5)), "Name      |  Age")
        __Check(CStr(row1Name.PadRight(10) & "|" & row1Age.PadLeft(5)), "Alice     |   30")
    End Sub
End Module
