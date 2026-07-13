use super::helpers::run_vb;

#[test]
fn xml_cdata_sections() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' XML literals support CDATA blocks
        Dim xml = <Data>
                      <![CDATA[Some <unescaped> data & characters]]>
                  </Data>
                  
        Console.WriteLine(xml.Value.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Some <unescaped> data & characters"]);
}
