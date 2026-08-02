' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_empty_marker_interface
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface ISerializableMarker
End Interface

Class DataPacket
    Implements ISerializableMarker
End Class

Module Program
    Sub Main()
        Dim p As Object = New DataPacket()
        __Check(CStr(TypeOf p Is ISerializableMarker), "True")
    End Sub
End Module
