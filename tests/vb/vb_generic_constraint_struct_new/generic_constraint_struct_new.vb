' vybe-test: vb/vb_generic_constraint_struct_new/generic_constraint_struct_new
' origin: languages/vb/tests/vb/test_vb_generic_constraint_struct_new.rs

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

' As Structure requires T to be a value type
Class ValueCache(Of T As Structure)
    Public Property Item As T
End Class

' As New requires T to have a parameterless constructor
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class Person
    Public Property Name As String = "Bob"
End Class

Module M
    Sub Main()
        Dim vc As New ValueCache(Of Integer)()
        vc.Item = 42
        __Check(CStr(vc.Item), "42")
        
        Dim f As New Factory(Of Person)()
        Dim p = f.Create()
        __Check(CStr(p.Name), "Bob")
    End Sub
End Module
