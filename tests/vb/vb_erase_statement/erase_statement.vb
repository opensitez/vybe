' vybe-test: vb/vb_erase_statement/erase_statement
' origin: languages/vb/tests/vb/test_vb_erase_statement.rs

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
        Dim fixedArray(2) As Integer
        Dim dynArray() As Integer = {1, 2, 3}
        
        ' Erase clears the array
        Erase fixedArray ' Reinitializes elements to default (0)
        Erase dynArray   ' Sets the reference to Nothing
        
        __Check(CStr(fixedArray(0)), "0")
        __Check(CStr(dynArray Is Nothing), "True")
    End Sub
End Module
