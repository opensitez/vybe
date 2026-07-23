use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XmlSerializer Serialization & Deserialization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_serializer_roundtrip() {
    let src = r#"
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
            Console.WriteLine(deserialized.Title & ":" & deserialized.Price)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VB Guide:49.99"]);
}
