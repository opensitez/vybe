use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Literals
// ═══════════════════════════════════════════════════════════

#[test]
fn xml_literal_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' XML literals are a first-class citizen in VB.NET
        Dim xml = <book>
                      <title>VB.NET Guide</title>
                      <author>John Doe</author>
                  </book>
                  
        Console.WriteLine(xml.<title>.Value)
        Console.WriteLine(xml.<author>.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["VB.NET Guide", "John Doe"]);
}

#[test]
fn xml_literal_embedded_expressions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim year As Integer = 2026
        Dim xml = <report year=<%= year %>>
                      <status>Complete</status>
                  </report>
                  
        Console.WriteLine(xml.@year)
        Console.WriteLine(xml.<status>.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2026", "Complete"]);
}
