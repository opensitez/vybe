' vybe-test: vb/vb_system_boxing_unboxing_matrix/boxing_unboxing_trycast_returns_nothing_on_mismatch
' origin: languages/vb/tests/vb/test_vb_system_boxing_unboxing_matrix.rs

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
        Dim boxed As Object = 12.5
        Dim intRef As Nullable(Of Integer)

        intRef = TryCast(boxed, Integer)
        __Check(CStr(intRef.HasValue), "False")

        Dim asObj As String = TryCast(boxed, String)
        __Check(CStr(asObj Is Nothing), "True")
    End Sub
End Module
