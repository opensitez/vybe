' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_enum_type_arg
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Enum State
    Off = 0
    OnVal = 1
End Enum

Structure StateWrapper(Of T As Structure)
    Public CurrentState As T
    Public Sub New(s As T)
        CurrentState = s
    End Sub
End Structure

Module Program
    Sub Main()
        Dim w As New StateWrapper(Of State)(State.OnVal)
        __Check(CStr(w.CurrentState.ToString()), "OnVal")
    End Sub
End Module
