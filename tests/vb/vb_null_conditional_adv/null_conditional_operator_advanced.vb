' vybe-test: vb/vb_null_conditional_adv/null_conditional_operator_advanced
' origin: languages/vb/tests/vb/test_vb_null_conditional_adv.rs

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

Class Node
    Public Property Value As String
    Public Property NextNode As Node
    Public Function GetName() As String
        Return Value
    End Function
End Class

Module M
    Sub Main()
        Dim root As New Node() With {.Value = "Root"}
        Dim empty As Node = Nothing
        
        ' Null conditional method call
        __Check(CStr(root?.GetName()), "Root")
        __Check(CStr(empty?.GetName() Is Nothing), "True")
        
        ' Null conditional indexing (if it was an array/list)
        Dim arr() As Integer = {1, 2, 3}
        Dim emptyArr() As Integer = Nothing
        
        __Check(CStr(arr?(0)), "1")
        ' We can't really print Nothing directly for integer in VB without it being 0 if not nullable, 
        ' but for arrays ?. indexing returns Nullable(Of T)
        __Check(CStr(emptyArr?(0).HasValue), "False")
    End Sub
End Module
