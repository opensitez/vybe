' vybe-test: vb/vb_structs_constructors/struct_constructor_initialization
' origin: languages/vb/tests/vb/test_vb_structs_constructors.rs

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

Structure Size
    Public Width As Integer
    Public Height As Integer
    
    ' Parameterized constructor
    Public Sub New(w As Integer, h As Integer)
        Width = w
        Height = h
    End Sub
End Structure

Module M
    Sub Main()
        Dim s As New Size(1024, 768)
        __Check(CStr(s.Width), "1024")
        __Check(CStr(s.Height), "768")
    End Sub
End Module
