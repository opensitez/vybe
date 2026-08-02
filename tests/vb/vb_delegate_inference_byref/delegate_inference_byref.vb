' vybe-test: vb/vb_delegate_inference_byref/delegate_inference_byref
' origin: languages/vb/tests/vb/test_vb_delegate_inference_byref.rs

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
    Delegate Sub ByRefAction(ByRef x As Integer)

    Sub Main()
        ' Delegate type inference with ByRef
        Dim act As ByRefAction = Sub(ByRef x As Integer) x += 1
        
        Dim val = 10
        act(val)
        __Check(CStr(val), "11")
    End Sub
End Module
