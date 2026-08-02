' vybe-test: vb/vb_array_resize_preserve_semantics/test_vb_array_redim_preserve_struct_array
' origin: languages/vb/tests/vb/test_vb_array_resize_preserve_semantics.rs

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

Structure Pair
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Dim pairs(0) As Pair
        pairs(0).X = 10 : pairs(0).Y = 20
        ReDim Preserve pairs(1)
        __Check(CStr(pairs(0).X & ":" & pairs(0).Y), "10:20")
        __Check(CStr(pairs(1).X & ":" & pairs(1).Y), "0:0")
    End Sub
End Module
