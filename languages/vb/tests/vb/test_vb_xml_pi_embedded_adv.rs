use super::helpers::run_vb;

#[test]
fn xml_pi_embedded_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim piData = "Version=1.0"
        ' Processing Instruction with embedded expression
        Dim xml = <?PI <%= piData %>?>
                  
        Console.WriteLine(xml.Data.Trim())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Version=1.0"]);
}
