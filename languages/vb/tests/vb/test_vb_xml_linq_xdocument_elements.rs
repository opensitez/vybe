use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XDocument & XElement Construction
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_xdocument_construction_linq_query() {
    let src = r#"
Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim items = {"A", "B", "C"}
        Dim doc As New XDocument(
            New XElement("root",
                From i In items Select <item><%= i %></item>
            )
        )
        Console.WriteLine(doc.Root.Elements("item").Count())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}
