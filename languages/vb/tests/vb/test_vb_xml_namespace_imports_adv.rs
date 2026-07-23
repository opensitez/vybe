use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Imports <xmlns:...> Global XML Namespaces
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_namespace_qualified_elements() {
    let src = r#"
Imports System.Xml.Linq
Imports <xmlns:ns="http://example.com/ns">

Module Program
    Sub Main()
        Dim elem As XElement = <ns:data ns:attr="val">Content</ns:data>
        Console.WriteLine(elem.Name.NamespaceName)
        Console.WriteLine(elem.@ns:attr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["http://example.com/ns", "val"]);
}
