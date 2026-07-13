use super::helpers::run_vb;

#[test]
fn imports_xml_namespace() {
    let out = run_vb(
        r#"
Imports <xmlns:ns="http://example.com/ns">

Module M
    Sub Main()
        Dim xml = <ns:Root>
                      <ns:Child>Value</ns:Child>
                  </ns:Root>
                  
        Console.WriteLine(xml.Name.NamespaceName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["http://example.com/ns"]);
}
