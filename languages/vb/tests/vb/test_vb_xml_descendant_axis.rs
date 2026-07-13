use super::helpers::run_vb;

#[test]
fn xml_descendant_axis() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <Root>
                      <Container>
                          <Item>One</Item>
                      </Container>
                      <Item>Two</Item>
                  </Root>
                  
        ' XML descendant axis operator ... returns all matching descendants at any level
        Dim items = xml...<Item>
        
        For Each item In items
            Console.WriteLine(item.Value)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["One", "Two"]);
}
