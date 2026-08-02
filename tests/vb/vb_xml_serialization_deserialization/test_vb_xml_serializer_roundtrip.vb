' vybe-test: vb/vb_xml_serialization_deserialization/test_vb_xml_serializer_roundtrip
' origin: languages/vb/tests/vb/test_vb_xml_serialization_deserialization.rs

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

Imports System.IO
Imports System.Xml.Serialization

Public Class BookDto
    Public Property Title As String
    Public Property Price As Double
End Class

Module Program
    Sub Main()
        Dim book As New BookDto() With {.Title = "VB Guide", .Price = 49.99}
        Dim serializer As New XmlSerializer(GetType(BookDto))

        Using ms As New MemoryStream()
            serializer.Serialize(ms, book)
            ms.Position = 0

            Dim deserialized As BookDto = CType(serializer.Deserialize(ms), BookDto)
            __Check(CStr(deserialized.Title & ":" & deserialized.Price), "VB Guide:49.99")
        End Using
    End Sub
End Module
