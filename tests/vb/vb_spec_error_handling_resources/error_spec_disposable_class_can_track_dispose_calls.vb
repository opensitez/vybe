' vybe-test: vb/vb_spec_error_handling_resources/error_spec_disposable_class_can_track_dispose_calls
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

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

Class Probe
    Implements IDisposable
    Public Shared Count As Integer
    Public Sub Dispose() Implements IDisposable.Dispose
        Count += 1
        __Check(CStr(Count), "1")
    End Sub
End Class
Module M
    Sub Main()
        Using value As New Probe()
        End Using
    End Sub
End Module
