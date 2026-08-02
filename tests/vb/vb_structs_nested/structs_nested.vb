' vybe-test: vb/vb_structs_nested/structs_nested
' origin: languages/vb/tests/vb/test_vb_structs_nested.rs

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

Structure Outer
    Public X As Integer
    
    Structure Inner
        Public Y As Integer
    End Structure
    
    Public InnerData As Inner
End Structure

Module M
    Sub Main()
        Dim o As New Outer()
        o.X = 10
        o.InnerData.Y = 20
        
        __Check(CStr(o.X), "10")
        __Check(CStr(o.InnerData.Y), "20")
    End Sub
End Module
