use super::helpers::run_vb;

#[test]
fn xml_literals_cdata() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        Dim xml = <Data><![CDATA[<Test> & "Quotes"]]></Data>
        Console.WriteLine(xml.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["<Test> & \"Quotes\""]);
}

#[test]
fn xml_literals_namespaces() {
    let out = run_vb(
        r#"
Imports System.Xml.Linq

Module M
    Sub Main()
        ' XML namespaces in literals
        Dim ns = <xml xmlns:ns1="http://example.com/ns1">
                     <ns1:Item>Value</ns1:Item>
                 </xml>
                 
        Console.WriteLine(ns.Elements().First().Name.LocalName)
        Console.WriteLine(ns.Elements().First().Name.NamespaceName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Item", "http://example.com/ns1"]);
}
