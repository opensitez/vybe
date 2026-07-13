use super::helpers::run_vb;

#[test]
fn xml_axis_properties_attributes() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' XML literals
        Dim book = <book id="123"><title>Programming in VB.NET</title><author>Jane Doe</author></book>
        
        ' Child axis property
        Console.WriteLine(book.<title>.Value)
        
        ' Attribute axis property
        Console.WriteLine(book.@id)
        
        ' XML interpolation
        Dim newTitle = "VB.NET Advanced"
        Dim book2 = <book><title><%= newTitle %></title></book>
        Console.WriteLine(book2.<title>.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Programming in VB.NET", "123", "VB.NET Advanced"]);
}
