//! `System.Xml.Linq`: XDocument, XElement, XAttribute, LINQ to XML queries.
use super::helpers::run_csharp;

#[test]
fn xelement_name_and_value_readable() {
    assert_eq!(
        run_csharp(
            r#"var el=new System.Xml.Linq.XElement("Item","hello");
Console.WriteLine(el.Name.LocalName); Console.WriteLine((string)el);"#
        ),
        &["Item", "hello"]
    );
}

#[test]
fn xdocument_root_element_accessible() {
    assert_eq!(
        run_csharp(
            r#"var doc=System.Xml.Linq.XDocument.Parse("<root><child>v</child></root>");
Console.WriteLine(doc.Root.Name.LocalName);"#
        ),
        &["root"]
    );
}

#[test]
fn xelement_attribute_read_back() {
    assert_eq!(
        run_csharp(
            r#"var el=new System.Xml.Linq.XElement("Node",
    new System.Xml.Linq.XAttribute("id","42"));
Console.WriteLine((string)el.Attribute("id"));"#
        ),
        &["42"]
    );
}

#[test]
fn xelement_descendants_query_returns_matching_nodes() {
    assert_eq!(
        run_csharp(
            r#"var doc=System.Xml.Linq.XDocument.Parse("<root><a>1</a><a>2</a></root>");
int count=doc.Root.Elements("a").Count();
Console.WriteLine(count);"#
        ),
        &["2"]
    );
}

#[test]
fn xelement_construction_with_children_builds_tree() {
    assert_eq!(
        run_csharp(
            r#"var xml=new System.Xml.Linq.XElement("Root",
    new System.Xml.Linq.XElement("Child","data"));
Console.WriteLine(xml.Element("Child").Value);"#
        ),
        &["data"]
    );
}
