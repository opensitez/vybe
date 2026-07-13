use super::helpers::run_vb;

#[test]
fn xml_axis_indexer() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim xml = <Root>
                      <Child>A</Child>
                      <Child>B</Child>
                  </Root>
                  
        ' XML child axis with indexer (0-based)
        Console.WriteLine(xml.<Child>(0).Value)
        Console.WriteLine(xml.<Child>(1).Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "B"]);
}
