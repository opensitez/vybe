use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Literals & Embedded Expressions <%= expr %>
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_literal_embedded_element_and_attribute() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim name As String = "Laptop"
        Dim price As Double = 999.99
        Dim id As Integer = 101

        Dim doc As XElement = <product id=<%= id %>>
                                  <name><%= name %></name>
                                  <price><%= price %></price>
                              </product>

        Console.WriteLine(doc.@id)
        Console.WriteLine(doc.<name>.Value)
        Console.WriteLine(doc.<price>.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["101", "Laptop", "999.99"]);
}
