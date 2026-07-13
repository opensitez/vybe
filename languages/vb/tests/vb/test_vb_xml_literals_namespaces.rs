use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Literals Advanced (Namespaces)
// ═══════════════════════════════════════════════════════════

#[test]
fn xml_literals_namespaces() {
    let out = run_vb(
        r#"
Imports <xmlns:ns="http://example.com/ns">

Module M
    Sub Main()
        Dim xml = <ns:Root>
                      <ns:Child>Value</ns:Child>
                  </ns:Root>
                  
        ' Need to use GetNamespace to query with namespaces
        Dim ns = GetXmlNamespace(ns)
        Console.WriteLine(xml.Element(ns + "Child").Value)
    End Sub
End Module
"#,
    );
    // As long as it parses the XML namespace syntax and compiles it's fine for VM baseline
    assert_eq!(out, vec!["Value"]);
}
