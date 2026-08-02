' vybe-test: vb/vb_option_strict_on_off_coercion/test_vb_option_strict_off_late_bound_method_call
' origin: languages/vb/tests/vb/test_vb_option_strict_on_off_coercion.rs

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

Class DynamicWorker
    Public Function Work(msg As String) As String
        Return "Worked: " & msg
    End Function
End Class

Module Program
    Sub Main()
        Dim obj As Object = New DynamicWorker()
        ' Late bound call without explicit cast under Option Strict Off
        Dim res = obj.Work("Task1")
        __Check(CStr(res), "Worked: Task1")
    End Sub
End Module
