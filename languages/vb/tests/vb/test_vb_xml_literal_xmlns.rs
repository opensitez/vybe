use super::helpers::run_vb;

#[test]
fn xml_literal_xmlns() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' XML literal with inline xmlns
        Dim xml = <Root xmlns:ns="http://test.com">
                      <ns:Child>Val</ns:Child>
                  </Root>
                  
        Console.WriteLine(xml.Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Root"]);
}
