' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_constructor_overloads
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

Structure FlexBox(Of T)
    Public Val As T
    Public Name As String
    Public Sub New(v As T)
        Val = v : Name = "Unnamed"
    End Sub
    Public Sub New(v As T, n As String)
        Val = v : Name = n
    End Sub
End Structure

Module Program
    Sub Main()
        Dim b1 As New FlexBox(Of Integer)(10)
        Dim b2 As New FlexBox(Of Integer)(20, "Custom")
        __Check(CStr(b1.Name & "|" & b2.Name), "Unnamed|Custom")
    End Sub
End Module
