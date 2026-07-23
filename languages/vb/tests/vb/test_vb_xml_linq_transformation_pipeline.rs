use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Xml.Linq (XDocument, XElement, XAttribute) Pipelines
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_xml_element_creation_and_value_retrieval() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim elem As New XElement("User", "Alice")
        Console.WriteLine(elem.Name.LocalName & "=" & elem.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["User=Alice"]);
}

#[test]
fn test_vb_xml_element_with_attributes() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim elem As New XElement("Product",
            New XAttribute("Id", "P100"),
            New XAttribute("Price", "29.99"),
            "Widget"
        )
        Console.WriteLine(elem.Attribute("Id").Value & "|" & elem.Attribute("Price").Value & "|" & elem.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["P100|29.99|Widget"]);
}

#[test]
fn test_vb_xml_document_parse_and_query_elements() {
    let src = r#"
Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim xmlStr = "<Catalog><Item Price='10'>A</Item><Item Price='20'>B</Item></Catalog>"
        Dim doc = XDocument.Parse(xmlStr)

        Dim items = doc.Root.Elements("Item").Select(Function(e) e.Value & ":" & e.Attribute("Price").Value)
        Console.WriteLine(String.Join(",", items))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A:10,B:20"]);
}

#[test]
fn test_vb_xml_linq_query_filtering_elements() {
    let src = r#"
Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim doc As New XDocument(
            New XElement("Orders",
                New XElement("Order", New XAttribute("Status", "Completed"), "O1"),
                New XElement("Order", New XAttribute("Status", "Pending"), "O2"),
                New XElement("Order", New XAttribute("Status", "Completed"), "O3")
            )
        )

        Dim completed = From o In doc.Root.Elements("Order")
                        Where o.Attribute("Status").Value = "Completed"
                        Select o.Value

        Console.WriteLine(String.Join(",", completed))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["O1,O3"]);
}

#[test]
fn test_vb_xml_element_add_remove_nodes() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim root As New XElement("List", New XElement("Item", "1"))
        root.Add(New XElement("Item", "2"))
        root.Element("Item").Remove() ' Removes first "1"

        Console.WriteLine(root.Element("Item").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_xml_namespaces_xnamespace_usage() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim ns As XNamespace = "http://example.com/ns"
        Dim elem As New XElement(ns + "Root",
            New XAttribute(XNamespace.Xmlns + "ex", ns),
            New XElement(ns + "Child", "Value")
        )
        Console.WriteLine(elem.Element(ns + "Child").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Value"]);
}

#[test]
fn test_vb_xml_descendants_search() {
    let src = r#"
Imports System.Linq
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim doc = XDocument.Parse("<Root><Group><Item>A</Item></Group><Item>B</Item></Root>")
        Dim allItems = doc.Descendants("Item").Select(Function(e) e.Value)
        Console.WriteLine(String.Join(",", allItems))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A,B"]);
}

#[test]
fn test_vb_xml_transformation_to_domain_objects() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml.Linq

Class Book
    Public Property Title As String
    Public Property Author As String
End Class

Module Program
    Sub Main()
        Dim xmlStr = "<Library><Book><Title>VB Guide</Title><Author>Alice</Author></Book></Library>"
        Dim doc = XDocument.Parse(xmlStr)

        Dim books = doc.Root.Elements("Book").Select(Function(b) New Book With {
            .Title = b.Element("Title").Value,
            .Author = b.Element("Author").Value
        }).ToList()

        Console.WriteLine(books(0).Title & " by " & books(0).Author)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VB Guide by Alice"]);
}

#[test]
fn test_vb_xml_domain_objects_to_xml_transformation() {
    let src = r#"
Imports System.Collections.Generic
Imports System.Linq
Imports System.Xml.Linq

Class Item
    Public Property Code As String
    Public Property Qty As Integer
End Class

Module Program
    Sub Main()
        Dim items As New List(Of Item) From {
            New Item With {.Code = "I1", .Qty = 10},
            New Item With {.Code = "I2", .Qty = 20}
        }

        Dim root As New XElement("Inventory",
            From i In items Select New XElement("Item", New XAttribute("Code", i.Code), i.Qty)
        )

        Console.WriteLine(root.Elements("Item").Count() & "|" & root.ToString().Contains("Code=""I1"""))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|True"]);
}

#[test]
fn test_vb_xml_replace_with_node_mutation() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim doc = XDocument.Parse("<Root><OldNode>Original</OldNode></Root>")
        doc.Root.Element("OldNode").ReplaceWith(New XElement("NewNode", "Replaced"))
        Console.WriteLine(doc.Root.FirstNode.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["<NewNode>Replaced</NewNode>"]);
}

#[test]
fn test_vb_xml_explicit_value_casting_to_primitives() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim elem = XElement.Parse("<Data Count='42' Ratio='3.14' Active='true'>Payload</Data>")

        Dim count As Integer = CInt(elem.Attribute("Count"))
        Dim ratio As Double = CDbl(elem.Attribute("Ratio"))
        Dim active As Boolean = CBool(elem.Attribute("Active"))

        Console.WriteLine(count & "|" & ratio & "|" & active)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42|3.14|True"]);
}

#[test]
fn test_vb_xml_null_attribute_cast_returns_nothing() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim elem As New XElement("Node")
        Dim missingAttr As XAttribute = elem.Attribute("Missing")
        Dim val As String = CStr(missingAttr)
        Console.WriteLine(val Is Nothing)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_xml_set_attribute_value_upsert() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim elem As New XElement("Setting")
        elem.SetAttributeValue("Key", "K1") ' Inserts
        elem.SetAttributeValue("Key", "K2") ' Updates
        Console.WriteLine(elem.Attribute("Key").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["K2"]);
}

#[test]
fn test_vb_xml_set_element_value_upsert() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim root As New XElement("Config")
        root.SetElementValue("Option", "Val1")
        root.SetElementValue("Option", "Val2")
        Console.WriteLine(root.Element("Option").Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Val2"]);
}

#[test]
fn test_vb_xml_ancestors_and_parent_navigation() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim doc = XDocument.Parse("<GrandParent><Parent><Child>Leaf</Child></Parent></GrandParent>")
        Dim child = doc.Descendants("Child").First()
        Dim parentName = child.Parent.Name.LocalName
        Dim rootName = child.Ancestors().Last().Name.LocalName
        Console.WriteLine(parentName & "|" & rootName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Parent|GrandParent"]);
}

#[test]
fn test_vb_xml_cdata_section_support() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim cdata As New XCData("<html><body>Test</body></html>")
        Dim elem As New XElement("Content", cdata)
        Console.WriteLine(elem.Value & "|" & (TypeOf elem.FirstNode Is XCData))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["<html><body>Test</body></html>|True"]);
}

#[test]
fn test_vb_xml_comments_and_node_types() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim doc As New XDocument(
            New XComment("Catalog Comment"),
            New XElement("Catalog")
        )
        Console.WriteLine(doc.FirstNode.NodeType.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Comment"]);
}

#[test]
fn test_vb_xml_save_and_load_memory_stream() {
    let src = r#"
Imports System.IO
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim origDoc As New XDocument(New XElement("Root", "StreamTest"))
        Using ms As New MemoryStream()
            origDoc.Save(ms)
            ms.Position = 0
            Dim restoredDoc = XDocument.Load(ms)
            Console.WriteLine(restoredDoc.Root.Value)
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["StreamTest"]);
}

#[test]
fn test_vb_xml_invalid_xml_parse_throws_xml_exception() {
    let src = r#"
Imports System.Xml
Imports System.Xml.Linq

Module Program
    Sub Main()
        Try
            XDocument.Parse("<UnclosedTag>Content")
        Catch ex As XmlException
            Console.WriteLine("XmlException Caught on Malformed XML")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["XmlException Caught on Malformed XML"]);
}

#[test]
fn test_vb_xml_deep_equals_structural_comparison() {
    let src = r#"
Imports System.Xml.Linq

Module Program
    Sub Main()
        Dim e1 = XElement.Parse("<Item Id='1'>A</Item>")
        Dim e2 = XElement.Parse("<Item Id='1'>A</Item>")
        Dim e3 = XElement.Parse("<Item Id='2'>A</Item>")

        Console.WriteLine(XNode.DeepEquals(e1, e2) & "|" & XNode.DeepEquals(e1, e3))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}
