' vybe-test: vb/vb_structs_byval_byref/struct_byref_parameter
' origin: languages/vb/tests/vb/test_vb_structs_byval_byref.rs

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

Structure Coordinate
    Public Lat As Double
    Public Lon As Double
End Structure

Module M
    Sub ModifyCoord(ByRef c As Coordinate)
        c.Lat = 99.9
    End Sub

    Sub Main()
        Dim loc As Coordinate
        loc.Lat = 45.0
        ModifyCoord(loc)
        ' Should change because it was passed by reference
        __Check(CStr(loc.Lat), "99.9")
    End Sub
End Module
