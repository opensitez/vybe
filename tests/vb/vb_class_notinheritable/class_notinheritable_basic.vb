' vybe-test: vb/vb_class_notinheritable/class_notinheritable_basic
' origin: languages/vb/tests/vb/test_vb_class_notinheritable.rs

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

NotInheritable Class MathUtils
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Module M
    Sub Main()
        Dim utils As New MathUtils()
        __Check(CStr(utils.Add(10, 20)), "30")
    End Sub
End Module
