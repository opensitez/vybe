use super::helpers::run_vb;

#[test]
fn xml_attribute_axis() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <Data id="42" name="Item" />
                  
        ' XML attribute axis operator @
        Console.WriteLine(xml.@id)
        Console.WriteLine(xml.@name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Item"]);
}
