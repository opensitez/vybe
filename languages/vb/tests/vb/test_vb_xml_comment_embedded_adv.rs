use super::helpers::run_vb;

#[test]
fn xml_comment_embedded_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim info = "TestComment"
        ' Embedded expressions inside XML comments
        Dim xml = <!-- <%= info %> -->
                  
        Console.WriteLine(xml.Value.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["TestComment"]);
}
