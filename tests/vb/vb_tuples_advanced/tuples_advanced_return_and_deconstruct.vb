' vybe-test: vb/vb_tuples_advanced/tuples_advanced_return_and_deconstruct
' origin: languages/vb/tests/vb/test_vb_tuples_advanced.rs

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

Module M
    Function GetCoordinates() As (X As Integer, Y As Integer)
        Return (10, 20)
    End Function

    Sub Main()
        ' Return named tuple
        Dim coords = GetCoordinates()
        __Check(CStr(coords.X), "10")
        __Check(CStr(coords.Y), "20")
        
        ' Deconstruction into existing variables (or new ones)
        Dim a, b As Integer
        (a, b) = GetCoordinates()
        __Check(CStr(a), "10")
        __Check(CStr(b), "20")
    End Sub
End Module
