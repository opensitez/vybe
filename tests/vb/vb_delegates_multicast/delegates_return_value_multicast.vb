' vybe-test: vb/vb_delegates_multicast/delegates_return_value_multicast
' origin: languages/vb/tests/vb/test_vb_delegates_multicast.rs

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

Delegate Function Calc(x As Integer) As Integer

Module M
    Function DoubleIt(x As Integer) As Integer
        Return x * 2
    End Function
    
    Function TripleIt(x As Integer) As Integer
        Return x * 3
    End Function

    Sub Main()
        Dim d1 As Calc = AddressOf DoubleIt
        Dim d2 As Calc = AddressOf TripleIt
        
        Dim d3 As Calc = CType([Delegate].Combine(d1, d2), Calc)
        
        ' When a multicast delegate returns a value, it returns the value from the last method invoked.
        Dim result As Integer = d3(5)
        __Check(CStr(result), "15")
    End Sub
End Module
