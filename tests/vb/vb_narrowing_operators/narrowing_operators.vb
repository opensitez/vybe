' vybe-test: vb/vb_narrowing_operators/narrowing_operators
' origin: languages/vb/tests/vb/test_vb_narrowing_operators.rs

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

Class Wrapper
    Public Value As Integer
    
    ' Narrowing explicitly requires casting
    Public Shared Narrowing Operator CType(w As Wrapper) As Integer
        Return w.Value
    End Operator
    
    ' Widening allows implicit casting
    Public Shared Widening Operator CType(i As Integer) As Wrapper
        Return New Wrapper() With {.Value = i}
    End Operator
End Class

Module M
    Sub Main()
        ' Widening implicit
        Dim w As Wrapper = 42
        __Check(CStr(w.Value), "42")
        
        ' Narrowing explicit
        Dim i As Integer = CType(w, Integer)
        __Check(CStr(i), "42")
    End Sub
End Module
