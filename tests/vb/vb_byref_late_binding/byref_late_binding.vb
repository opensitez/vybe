' vybe-test: vb/vb_byref_late_binding/byref_late_binding
' origin: languages/vb/tests/vb/test_vb_byref_late_binding.rs

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

Option Strict Off

Module M
    Sub ModifyRef(ByRef val As Integer)
        val += 10
    End Sub

    Sub Main()
        Dim obj As Object = 5
        ' Late binding ByRef passes the object, which is unboxed, modified, and re-boxed
        ' Wait, if we pass an Object to an Integer ByRef parameter with Option Strict Off,
        ' VB.NET generates a temporary variable and copies it back.
        ModifyRef(obj)
        __Check(CStr(obj), "15")
    End Sub
End Module
