use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Literals Advanced (CDATA)
// ═══════════════════════════════════════════════════════════

#[test]
fn xml_literals_cdata() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <root>
                      <![CDATA[ <html><body>Hello!</body></html> ]]>
                  </root>
                  
        ' CDATA section preserves all characters
        Console.WriteLine(xml.Value.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["<html><body>Hello!</body></html>"]);
}
