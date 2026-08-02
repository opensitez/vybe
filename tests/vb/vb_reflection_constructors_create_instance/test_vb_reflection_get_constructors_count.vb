' vybe-test: vb/vb_reflection_constructors_create_instance/test_vb_reflection_get_constructors_count
' origin: languages/vb/tests/vb/test_vb_reflection_constructors_create_instance.rs

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

Class MultiCtor
    Public Sub New() : End Sub
    Public Sub New(a As Integer) : End Sub
    Public Sub New(a As Integer, b As String) : End Sub
End Class

Module Program
    Sub Main()
        Dim ctors = GetType(MultiCtor).GetConstructors()
        __Check(CStr(ctors.Length), "3")
    End Sub
End Module
