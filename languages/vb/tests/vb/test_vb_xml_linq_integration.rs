use super::helpers::run_vb;

#[test]
fn xml_linq_integration() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim items = {1, 2, 3}
        
        ' Embedded expression with LINQ inside XML literal
        Dim xml = <Root>
                      <%= From x In items Select <Item><%= x %></Item> %>
                  </Root>
                  
        Console.WriteLine(xml.<Item>.Count())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}
