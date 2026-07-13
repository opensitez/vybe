use super::helpers::run_vb;

#[test]
fn xml_nested_literals() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim tag = "Inner"
        ' Deeply nested XML literals with embedded expressions
        Dim xml = <Outer>
                      <Middle>
                          <<%= tag %> id="1" />
                      </Middle>
                  </Outer>
                  
        Console.WriteLine(xml.<Middle>.Elements()(0).Name.LocalName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Inner"]);
}
