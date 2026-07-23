use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Axis Properties (..., .@, .<>)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_descendants_axis_property() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim catalog = <catalog>
                          <book id="1">
                              <title>Book One</title>
                          </book>
                          <book id="2">
                              <title>Book Two</title>
                          </book>
                      </catalog>

        Dim titles = catalog...<title>
        For Each t In titles
            Console.WriteLine(t.Value)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Book One", "Book Two"]);
}
