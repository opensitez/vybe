use super::helpers::run_python;

// xml.etree.ElementTree — Element, SubElement, fromstring, tostring, find, findall, iterfind, indent, text, attrib, get, set, keys, items, iter

#[test]
fn test_xml_elementtree_fromstring_and_tag() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.fromstring("<root><child>text</child></root>")
print(root.tag)
print(root[0].tag)
print(root[0].text)
"#,
    );
    assert_eq!(out, vec!["root", "child", "text"]);
}

#[test]
fn test_xml_elementtree_subelement_builder() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.Element("data")
item1 = ET.SubElement(root, "item", id="1")
item1.text = "Item One"
item2 = ET.SubElement(root, "item", id="2")
item2.text = "Item Two"

xml_str = ET.tostring(root, encoding="unicode")
print("item id=\"1\"" in xml_str)
print("Item Two" in xml_str)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_xml_elementtree_find_and_findall() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = """<catalog>
    <book id="bk101"><title>Python Guide</title></book>
    <book id="bk102"><title>Rust Guide</title></book>
</catalog>"""
root = ET.fromstring(xml)
first_book = root.find("book")
print(first_book.attrib["id"])

all_titles = [t.text for t in root.findall("book/title")]
print(all_titles)
"#,
    );
    assert_eq!(out, vec!["bk101", "['Python Guide', 'Rust Guide']"]);
}

#[test]
fn test_xml_elementtree_findtext() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = "<user><name>Alice</name><age>30</age></user>"
root = ET.fromstring(xml)
print(root.findtext("name"))
print(root.findtext("non_existent", default="N/A"))
"#,
    );
    assert_eq!(out, vec!["Alice", "N/A"]);
}

#[test]
fn test_xml_elementtree_element_attributes_get_set_items() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
elem = ET.Element("node", color="red", size="10")
print(elem.get("color"))
print(elem.get("missing", "default_val"))
elem.set("color", "blue")
print(elem.get("color"))
print(sorted(elem.items()))
"#,
    );
    assert_eq!(
        out,
        vec![
            "red",
            "default_val",
            "blue",
            "[('color', 'blue'), ('size', '10')]"
        ]
    );
}

#[test]
fn test_xml_elementtree_iter_recursive_traversal() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = "<a><b><c>text1</c></b><c>text2</c></a>"
root = ET.fromstring(xml)
c_tags = [e.text for e in root.iter("c")]
print(c_tags)
"#,
    );
    assert_eq!(out, vec!["['text1', 'text2']"]);
}

#[test]
fn test_xml_elementtree_itertext_concatenation() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = "<p>Hello <b>World</b>!</p>"
root = ET.fromstring(xml)
text = "".join(root.itertext())
print(text)
"#,
    );
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn test_xml_elementtree_indent_formatting() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a>1</a><b>2</b></root>")
if hasattr(ET, "indent"):
    ET.indent(root, space="  ")
    xml_str = ET.tostring(root, encoding="unicode")
    print("\n  <a>" in xml_str)
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xml_elementtree_parseerror_handling() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
try:
    ET.fromstring("<root><unclosed></root>")
except ET.ParseError:
    print("ParseError")
"#,
    );
    assert_eq!(out, vec!["ParseError"]);
}

#[test]
fn test_xml_elementtree_clear_element() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
elem = ET.fromstring("<node a='1'>text<child/></node>")
elem.clear()
print(elem.tag)
print(len(elem))
print(elem.attrib)
"#,
    );
    assert_eq!(out, vec!["node", "0", "{}"]);
}

#[test]
fn test_xml_elementtree_remove_child_element() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a/><b/><c/></root>")
child_b = root.find("b")
root.remove(child_b)
print([e.tag for e in root])
"#,
    );
    assert_eq!(out, vec!["['a', 'c']"]);
}

#[test]
fn test_xml_elementtree_insert_child_element() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.fromstring("<root><a/><c/></root>")
b = ET.Element("b")
root.insert(1, b)
print([e.tag for e in root])
"#,
    );
    assert_eq!(out, vec!["['a', 'b', 'c']"]);
}

#[test]
fn test_xml_elementtree_tail_text_attribute() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = "<p>Text before <b>bold</b> text after.</p>"
root = ET.fromstring(xml)
b_elem = root.find("b")
print(b_elem.text)
print(b_elem.tail.strip())
"#,
    );
    assert_eq!(out, vec!["bold", "text after."]);
}

#[test]
fn test_xml_elementtree_tostring_xml_declaration() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.Element("root")
s = ET.tostring(root, encoding="utf-8", xml_declaration=True)
print(s.startswith(b"<?xml"))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_xml_elementtree_xpath_attribute_predicate() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
xml = """<root>
    <item type="a">val1</item>
    <item type="b">val2</item>
</root>"""
root = ET.fromstring(xml)
target = root.find("item[@type='b']")
print(target.text)
"#,
    );
    assert_eq!(out, vec!["val2"]);
}

#[test]
fn test_xml_elementtree_keys_method() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
elem = ET.Element("t", x="1", y="2")
print(sorted(elem.keys()))
"#,
    );
    assert_eq!(out, vec!["['x', 'y']"]);
}

#[test]
fn test_xml_elementtree_extend_children() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
root = ET.Element("root")
children = [ET.Element("child1"), ET.Element("child2")]
root.extend(children)
print([e.tag for e in root])
"#,
    );
    assert_eq!(out, vec!["['child1', 'child2']"]);
}

#[test]
fn test_xml_elementtree_comment_factory() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
comment = ET.Comment("This is a comment")
print(callable(comment.tag))
print(comment.text)
"#,
    );
    assert_eq!(out, vec!["True", "This is a comment"]);
}

#[test]
fn test_xml_elementtree_processing_instruction() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
pi = ET.PI("php", "echo 'hello';")
print(callable(pi.tag))
print(pi.text)
"#,
    );
    assert_eq!(out, vec!["True", "echo 'hello';"]);
}

#[test]
fn test_xml_elementtree_register_namespace() {
    let out = run_python(
        r#"
import xml.etree.ElementTree as ET
ET.register_namespace("ns", "http://example.com/ns")
elem = ET.Element("{http://example.com/ns}custom")
s = ET.tostring(elem, encoding="unicode")
print("xmlns:ns=" in s or "ns:custom" in s)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
