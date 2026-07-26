use super::helpers::run_pascal;

// ═══════════════════════════════════════════════════════════
// Category 70: XML Document Processing & DOM Nodes (IXMLDocument)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_xmldocument_loadfromxml_rootnode() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><title>PascalXML</title></root>');
  WriteLn(doc.DocumentElement.NodeName);
end.
"#,
    );
    assert_eq!(out, vec!["root"]);
}

#[test]
fn test_xmldocument_read_child_node_text() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><item>Widget</item></root>');
  WriteLn(doc.DocumentElement.ChildNodes['item'].Text);
end.
"#,
    );
    assert_eq!(out, vec!["Widget"]);
}

#[test]
fn test_xmldocument_read_attribute() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><user id="55">Alice</user></root>');
  WriteLn(doc.DocumentElement.ChildNodes['user'].Attributes['id']);
end.
"#,
    );
    assert_eq!(out, vec!["55"]);
}

#[test]
fn test_xmldocument_create_new_document() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; rootNode, childNode: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.Active := True;
  rootNode := doc.AddChild('config');
  childNode := rootNode.AddChild('port');
  childNode.Text := '8080';
  WriteLn(rootNode.ChildNodes['port'].Text);
end.
"#,
    );
    assert_eq!(out, vec!["8080"]);
}

#[test]
fn test_xmldocument_set_attribute() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; node: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.Active := True;
  node := doc.AddChild('item');
  node.Attributes['status'] := 'active';
  WriteLn(node.Attributes['status']);
end.
"#,
    );
    assert_eq!(out, vec!["active"]);
}

#[test]
fn test_xmldocument_savetoxml_output() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; xmlStr: String;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<data><val>10</val></data>');
  doc.SaveToXML(xmlStr);
  WriteLn(Pos('<val>10</val>', xmlStr) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xmldocument_findnode_by_name() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; node: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><a/><b/><c/></root>');
  node := doc.DocumentElement.ChildNodes.FindNode('b');
  WriteLn(node <> nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xmldocument_childnodes_count_and_iteration() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; i: Integer;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><item1/><item2/></root>');
  WriteLn(doc.DocumentElement.ChildNodes.Count);
  for i := 0 to doc.DocumentElement.ChildNodes.Count - 1 do
    WriteLn(doc.DocumentElement.ChildNodes[i].NodeName);
end.
"#,
    );
    assert_eq!(out, vec!["2", "item1", "item2"]);
}

#[test]
fn test_xmldocument_remove_child_node() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; rootNode, childNode: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><keep/><remove/></root>');
  rootNode := doc.DocumentElement;
  childNode := rootNode.ChildNodes.FindNode('remove');
  rootNode.ChildNodes.Remove(childNode);
  WriteLn(rootNode.ChildNodes.Count);
end.
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_xmldocument_nested_tree_traversal() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; innerNode: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<level1><level2><level3>DeepValue</level3></level2></level1>');
  innerNode := doc.DocumentElement.ChildNodes['level2'].ChildNodes['level3'];
  WriteLn(innerNode.Text);
end.
"#,
    );
    assert_eq!(out, vec!["DeepValue"]);
}

#[test]
fn test_xmldocument_attribute_count_and_query() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; node: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<element k1="v1" k2="v2"/>');
  node := doc.DocumentElement;
  WriteLn(node.AttributeNodes.Count);
  WriteLn(node.Attributes['k1']);
  WriteLn(node.Attributes['k2']);
end.
"#,
    );
    assert_eq!(out, vec!["2", "v1", "v2"]);
}

#[test]
fn test_xmldocument_interface_reference_counting() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
procedure ProcessXML(doc: IXMLDocument);
begin
  WriteLn(doc.DocumentElement.NodeName);
end;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<refcount_test/>');
  ProcessXML(doc);
end.
"#,
    );
    assert_eq!(out, vec!["refcount_test"]);
}

#[test]
fn test_xmldocument_node_type_element() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<testnode/>');
  WriteLn(doc.DocumentElement.NodeType = ntElement);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xmldocument_empty_element_text() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><empty/></root>');
  WriteLn(Length(doc.DocumentElement.ChildNodes['empty'].Text));
end.
"#,
    );
    assert_eq!(out, vec!["0"]);
}

#[test]
fn test_xmldocument_modify_existing_node_text() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; node: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><title>OldTitle</title></root>');
  node := doc.DocumentElement.ChildNodes['title'];
  node.Text := 'NewTitle';
  WriteLn(node.Text);
end.
"#,
    );
    assert_eq!(out, vec!["NewTitle"]);
}

#[test]
fn test_xmldocument_has_child_nodes_check() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><child/></root>');
  WriteLn(doc.DocumentElement.HasChildNodes);
  WriteLn(doc.DocumentElement.ChildNodes['child'].HasChildNodes);
end.
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_xmldocument_add_multiple_children() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; root: IXMLNode; i: Integer;
begin
  doc := TXMLDocument.Create(nil);
  doc.Active := True;
  root := doc.AddChild('items');
  for i := 1 to 3 do
    root.AddChild('item').Text := i.ToString;
  WriteLn(root.ChildNodes.Count);
end.
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_xmldocument_xml_declaration_header() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; xmlStr: String;
begin
  doc := TXMLDocument.Create(nil);
  doc.Active := True;
  doc.Version := '1.0';
  doc.Encoding := 'UTF-8';
  doc.AddChild('root');
  doc.SaveToXML(xmlStr);
  WriteLn(Pos('version="1.0"', xmlStr) > 0);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xmldocument_clone_node() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument; origNode, clonedNode: IXMLNode;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><item id="1">Original</item></root>');
  origNode := doc.DocumentElement.ChildNodes['item'];
  clonedNode := origNode.CloneNode(True);
  WriteLn(clonedNode.Text);
  WriteLn(clonedNode.Attributes['id']);
end.
"#,
    );
    assert_eq!(out, vec!["Original", "1"]);
}

#[test]
fn test_xmldocument_clear_document() {
    let out = run_pascal(
        r#"
program Test;
uses XmlIntf, XmlDoc;
var doc: IXMLDocument;
begin
  doc := TXMLDocument.Create(nil);
  doc.LoadFromXML('<root><data/></root>');
  doc.Active := False;
  WriteLn(doc.DocumentElement = nil);
end.
"#,
    );
    assert_eq!(out, vec!["True"]);
}
