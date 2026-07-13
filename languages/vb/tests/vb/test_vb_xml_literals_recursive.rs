use super::helpers::run_vb;

#[test]
fn xml_literals_recursive() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim name = "World"
        Dim inner = <Inner>Hello <%= name %></Inner>
        Dim outer = <Outer>
                        <%= inner %>
                    </Outer>
                    
        Console.WriteLine(outer.<Inner>.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}
