use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: XML Literals with Embedded Expressions
// ═══════════════════════════════════════════════════════════

#[test]
fn xml_literals_expressions() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim name As String = "Bob"
        Dim age As Integer = 30
        
        ' Embedded expressions in XML literals use <%= expr %>
        Dim userXml As XElement = 
            <User>
                <Name><%= name %></Name>
                <Age><%= age %></Age>
            </User>
            
        Console.WriteLine(userXml.<Name>.Value)
        Console.WriteLine(userXml.<Age>.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Bob", "30"]);
}
