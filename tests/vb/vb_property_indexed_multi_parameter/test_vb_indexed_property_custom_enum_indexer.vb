' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_custom_enum_indexer
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Enum Channel
    Red = 0
    Green = 1
    Blue = 2
End Enum

Class ColorPixel
    Private channels(2) As Byte
    Default Public Property Component(ch As Channel) As Byte
        Get
            Return channels(CInt(ch))
        End Get
        Set(value As Byte)
            channels(CInt(ch)) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New ColorPixel()
        p(Channel.Green) = 255
        __Check(CStr(p(Channel.Green)), "255")
    End Sub
End Module
