use super::helpers::run_vb;

#[test]
fn xml_embedded_attribute_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim val = "CustomValue"
        
        ' XML attribute value via embedded expression
        Dim xml = <Data id=<%= val %> />
                  
        Console.WriteLine(xml.@id)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["CustomValue"]);
}
