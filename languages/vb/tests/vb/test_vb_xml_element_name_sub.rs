use super::helpers::run_vb;

#[test]
fn xml_element_name_sub() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim name = "DynamicName"
        ' Element name substitution
        Dim xml = <<%= name %>>Content</>
                  
        Console.WriteLine(xml.Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["DynamicName"]);
}
