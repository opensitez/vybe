Module TestXml
    Dim passed As Integer = 0
    Dim failed As Integer = 0

    Sub Assert(condition As Boolean, label As String)
        If condition Then
            Console.WriteLine("  PASS: " & label)
            passed = passed + 1
        Else
            Console.WriteLine("  FAIL: " & label)
            failed = failed + 1
        End If
    End Sub

    Sub Main()
        Console.WriteLine("=== XML Full Test Suite ===")
        Console.WriteLine()

        ' ── 1. XDocument.Parse basic ────────────────────────────────
        Console.WriteLine("--- 1. XDocument.Parse ---")
        Dim xml As String = "<root><item id=""1"">Hello</item><item id=""2"">World</item></root>"
        Dim doc As Object = XDocument.Parse(xml)
        Assert(doc IsNot Nothing, "XDocument.Parse returns object")

        Dim root As Object = doc.Root
        Assert(root IsNot Nothing, "doc.Root exists")
        Assert(root.Name = "root", "root.Name = 'root'")
        Console.WriteLine()

        ' ── 2. Element and Elements queries ─────────────────────────
        Console.WriteLine("--- 2. Element / Elements ---")
        Dim firstItem As Object = root.Element("item")
        Assert(firstItem IsNot Nothing, "root.Element('item') found")
        Assert(firstItem.Name = "item", "firstItem.Name = 'item'")
        Assert(firstItem.Value = "Hello", "firstItem.Value = 'Hello'")

        Dim items As Object = root.Elements("item")
        Assert(items.Count = 2, "root.Elements('item').Count = 2")

        Dim allChildren As Object = root.Elements()
        Assert(allChildren.Count = 2, "root.Elements().Count = 2")
        Console.WriteLine()

        ' ── 3. Attribute access ─────────────────────────────────────
        Console.WriteLine("--- 3. Attributes ---")
        Dim attr As Object = firstItem.Attribute("id")
        Assert(attr IsNot Nothing, "firstItem.Attribute('id') found")
        Assert(attr.Name = "id", "attr.Name = 'id'")
        Assert(attr.Value = "1", "attr.Value = '1'")
        Assert(firstItem.HasAttributes = True, "firstItem.HasAttributes = True")
        Console.WriteLine()

        ' ── 4. Descendants ──────────────────────────────────────────
        Console.WriteLine("--- 4. Descendants ---")
        Dim xml2 As String = "<store><dept><item>A</item><item>B</item></dept><dept><item>C</item></dept></store>"
        Dim doc2 As Object = XDocument.Parse(xml2)
        Dim allItems As Object = doc2.Root.Descendants("item")
        Assert(allItems.Count = 3, "Descendants('item').Count = 3")

        Dim allDesc As Object = doc2.Root.Descendants()
        Assert(allDesc.Count = 5, "Descendants() total = 5 (2 dept + 3 item)")
        Console.WriteLine()

        ' ── 5. Constructor: New XElement ─────────────────────────────
        Console.WriteLine("--- 5. New XElement ---")
        Dim elem As Object = New XElement("person", "John")
        Assert(elem.Name = "person", "New XElement name = 'person'")
        Assert(elem.Value = "John", "New XElement value = 'John'")

        ' XElement with attribute
        Dim elem2 As Object = New XElement("book", New XAttribute("isbn", "123"), "Title Here")
        Assert(elem2.Name = "book", "book element name")
        Assert(elem2.HasAttributes = True, "book has attributes")
        Dim isbn As Object = elem2.Attribute("isbn")
        Assert(isbn.Value = "123", "isbn attribute value")
        Console.WriteLine()

        ' ── 6. Constructor: New XDocument ────────────────────────────
        Console.WriteLine("--- 6. New XDocument ---")
        Dim newDoc As Object = New XDocument(New XElement("catalog", New XElement("book", "VB Guide")))
        Dim newRoot As Object = newDoc.Root
        Assert(newRoot.Name = "catalog", "new doc root = 'catalog'")
        Dim book As Object = newRoot.Element("book")
        Assert(book.Value = "VB Guide", "book value = 'VB Guide'")
        Console.WriteLine()

        ' ── 7. Add / AddFirst ────────────────────────────────────────
        Console.WriteLine("--- 7. Add / AddFirst ---")
        Dim parent As Object = New XElement("list")
        parent.Add(New XElement("item", "First"))
        parent.Add(New XElement("item", "Second"))
        Dim kids As Object = parent.Elements()
        Assert(kids.Count = 2, "After Add: 2 children")

        parent.AddFirst(New XElement("item", "Zero"))
        Dim kids2 As Object = parent.Elements()
        Assert(kids2.Count = 3, "After AddFirst: 3 children")
        ' The first child should be "Zero"
        Dim first As Object = kids2(0)
        Assert(first.Value = "Zero", "AddFirst placed 'Zero' at index 0")
        Console.WriteLine()

        ' ── 8. SetValue ──────────────────────────────────────────────
        Console.WriteLine("--- 8. SetValue ---")
        Dim el As Object = New XElement("name", "OldValue")
        Assert(el.Value = "OldValue", "Before SetValue")
        el.SetValue("NewValue")
        Assert(el.Value = "NewValue", "After SetValue")
        Console.WriteLine()

        ' ── 9. SetAttributeValue ─────────────────────────────────────
        Console.WriteLine("--- 9. SetAttributeValue ---")
        Dim el2 As Object = New XElement("tag")
        el2.SetAttributeValue("color", "red")
        Dim colorAttr As Object = el2.Attribute("color")
        Assert(colorAttr.Value = "red", "SetAttributeValue adds new attr")

        ' Update existing
        el2.SetAttributeValue("color", "blue")
        Dim colorAttr2 As Object = el2.Attribute("color")
        Assert(colorAttr2.Value = "blue", "SetAttributeValue updates existing")
        Console.WriteLine()

        ' ── 10. SetElementValue ──────────────────────────────────────
        Console.WriteLine("--- 10. SetElementValue ---")
        Dim container As Object = New XElement("data")
        container.SetElementValue("name", "Alice")
        Dim nameEl As Object = container.Element("name")
        Assert(nameEl.Value = "Alice", "SetElementValue adds child")

        ' Update
        container.SetElementValue("name", "Bob")
        Dim nameEl2 As Object = container.Element("name")
        Assert(nameEl2.Value = "Bob", "SetElementValue updates child")
        Console.WriteLine()

        ' ── 11. RemoveAll / RemoveNodes / RemoveAttributes ───────────
        Console.WriteLine("--- 11. Remove methods ---")
        Dim el3 As Object = New XElement("test", New XAttribute("a", "1"), New XElement("child", "val"))
        Assert(el3.HasAttributes = True, "Before RemoveAttributes: has attrs")
        el3.RemoveAttributes()
        Assert(el3.HasAttributes = False, "After RemoveAttributes: no attrs")

        Assert(el3.HasElements = True, "Before RemoveNodes: has elements")
        el3.RemoveNodes()
        Assert(el3.HasElements = False, "After RemoveNodes: no elements")
        Console.WriteLine()

        ' ── 12. Properties: IsEmpty, HasElements, FirstNode, LastNode ─
        Console.WriteLine("--- 12. Properties ---")
        Dim emptyEl As Object = New XElement("empty")
        Assert(emptyEl.IsEmpty = True, "Empty element IsEmpty = True")

        emptyEl.Add(New XElement("child1"))
        emptyEl.Add(New XElement("child2"))
        Assert(emptyEl.IsEmpty = False, "Non-empty IsEmpty = False")
        Assert(emptyEl.HasElements = True, "HasElements = True")

        Dim fn As Object = emptyEl.FirstNode
        Assert(fn.Name = "child1", "FirstNode = child1")
        Dim ln As Object = emptyEl.LastNode
        Assert(ln.Name = "child2", "LastNode = child2")
        Console.WriteLine()

        ' ── 13. ToString serialization ───────────────────────────────
        Console.WriteLine("--- 13. ToString ---")
        Dim simple As Object = New XElement("greeting", "Hello World")
        Dim s As String = simple.ToString()
        Assert(s.Contains("<greeting>"), "ToString contains <greeting>")
        Assert(s.Contains("Hello World"), "ToString contains value")
        Assert(s.Contains("</greeting>"), "ToString contains </greeting>")
        Console.WriteLine()

        ' ── 14. XDocument ToString with declaration ──────────────────
        Console.WriteLine("--- 14. XDocument ToString ---")
        Dim fullDoc As Object = New XDocument( _
            New XDeclaration("1.0", "utf-8", "yes"), _
            New XElement("root", New XElement("child", "data")))
        Dim docStr As String = fullDoc.ToString()
        Assert(docStr.Contains("<?xml"), "Doc ToString has XML declaration")
        Assert(docStr.Contains("<root>"), "Doc ToString has root element")
        Console.WriteLine()

        ' ── 15. XComment ─────────────────────────────────────────────
        Console.WriteLine("--- 15. XComment ---")
        Dim comment As Object = New XComment("This is a comment")
        Assert(comment.Value = "This is a comment", "XComment.Value correct")
        Dim commentStr As String = comment.ToString()
        Assert(commentStr.Contains("<!--"), "XComment serializes with <!--")
        Console.WriteLine()

        ' ── 16. XDeclaration ─────────────────────────────────────────
        Console.WriteLine("--- 16. XDeclaration ---")
        Dim decl As Object = New XDeclaration("1.0", "utf-16", "no")
        Dim declStr As String = decl.ToString()
        Assert(declStr.Contains("utf-16"), "XDeclaration has encoding")
        Assert(declStr.Contains("standalone=""no"""), "XDeclaration has standalone")
        Console.WriteLine()

        ' ── 17. Parse with declaration ───────────────────────────────
        Console.WriteLine("--- 17. Parse with declaration ---")
        Dim xmlWithDecl As String = "<?xml version=""1.0"" encoding=""utf-8""?><data><value>42</value></data>"
        Dim docDecl As Object = XDocument.Parse(xmlWithDecl)
        Dim declObj As Object = docDecl.Declaration
        Assert(declObj IsNot Nothing, "Parsed declaration exists")
        Dim dataRoot As Object = docDecl.Root
        Assert(dataRoot.Name = "data", "Root after declaration = 'data'")
        Dim valEl As Object = dataRoot.Element("value")
        Assert(valEl.Value = "42", "Nested value = '42'")
        Console.WriteLine()

        ' ── 18. Nested element construction ──────────────────────────
        Console.WriteLine("--- 18. Nested construction ---")
        Dim html As Object = New XElement("html", _
            New XElement("head", New XElement("title", "My Page")), _
            New XElement("body", _
                New XElement("h1", "Welcome"), _
                New XElement("p", "Hello World")))
        Dim title As Object = html.Element("head").Element("title")
        Assert(title.Value = "My Page", "Nested: title = 'My Page'")
        Dim body As Object = html.Element("body")
        Dim paras As Object = body.Elements("p")
        Assert(paras.Count = 1, "Nested: 1 paragraph")
        Console.WriteLine()

        ' ── 19. Nodes method ─────────────────────────────────────────
        Console.WriteLine("--- 19. Nodes ---")
        Dim mixed As Object = New XElement("mixed")
        mixed.Add(New XElement("a", "1"))
        mixed.Add(New XComment("note"))
        mixed.Add(New XElement("b", "2"))
        Dim nodes As Object = mixed.Nodes()
        Assert(nodes.Count = 3, "Nodes() returns 3 items")
        Console.WriteLine()

        ' ── 20. Chained property access ──────────────────────────────
        Console.WriteLine("--- 20. Chained access ---")
        Dim xml3 As String = "<company><dept name=""Engineering""><employee>Alice</employee><employee>Bob</employee></dept></company>"
        Dim doc3 As Object = XDocument.Parse(xml3)
        Dim deptName As Object = doc3.Root.Element("dept").Attribute("name")
        Assert(deptName.Value = "Engineering", "Chained: dept name = Engineering")
        Dim emps As Object = doc3.Root.Element("dept").Elements("employee")
        Assert(emps.Count = 2, "Chained: 2 employees")
        Console.WriteLine()

        ' ── 21. Self-closing elements ────────────────────────────────
        Console.WriteLine("--- 21. Self-closing elements ---")
        Dim xml4 As String = "<form><input type=""text"" name=""first"" /><input type=""submit"" /></form>"
        Dim doc4 As Object = XDocument.Parse(xml4)
        Dim inputs As Object = doc4.Root.Elements("input")
        Assert(inputs.Count = 2, "Self-closing: 2 inputs")
        Dim firstInput As Object = inputs(0)
        Assert(firstInput.Attribute("type").Value = "text", "First input type = text")
        Assert(firstInput.Attribute("name").Value = "first", "First input name = first")
        Console.WriteLine()

        ' ── Summary ──────────────────────────────────────────────────
        Console.WriteLine()
        Console.WriteLine("=== XML Results: " & passed & " passed, " & failed & " failed ===")
        If failed = 0 Then
            Console.WriteLine("=== ALL XML TESTS PASSED ===")
        End If
    End Sub
End Module
